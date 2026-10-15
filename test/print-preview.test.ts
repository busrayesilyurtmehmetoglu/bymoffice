import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_SHEET_PAGES } from '../src/layout/sheet-pages';
import { createSheet, emptyStyleSheet } from '../src/model/sheet/types';
import type { Cell, Row, Sheet, SheetPageSetup, Workbook } from '../src/model/sheet/types';
import { DocumentPrintError } from '../src/ui/doc-print';
import {
  CanvasPool,
  MAX_PREVIEW_ZOOM,
  MIN_PREVIEW_ZOOM,
  PREVIEW_GAP,
  PREVIEW_ZOOM_STEPS,
  SettleClock,
  clampZoom,
  fitZoom,
  maxPreviewZoom,
  previewWindow,
  layoutPreview,
  pageAtScroll,
  previewBudget,
  previewPixelWidths,
  previewTransition,
  scrollForPage,
  stepZoom,
  visiblePages,
  type PreviewEvent,
  type PreviewState,
} from '../src/ui/preview-layout';
import { jobFileName, jobSheets, planSheetPages, previewSheets } from '../src/ui/sheet-print';
import { ImageStore } from '../src/render/images';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

const A4 = { width: 794, height: 1123 };
const A4_LANDSCAPE = { width: 1123, height: 794 };

function row(cells: Cell[]): Row {
  return { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells };
}

function makeSheet(name: string, rows: number, setup?: SheetPageSetup, state: Sheet['state'] = 'visible'): Sheet {
  const sheet = createSheet({ name, sheetId: 1, state, kind: 'worksheet', part: `/xl/worksheets/${name}.xml` });
  for (let r = 0; r < rows; r++) sheet.rows.set(r, row([{ kind: 'number', value: r, style: 0 }, { kind: 'string', value: `r${r}`, style: 0 }]));
  if (setup) sheet.pageSetup = setup;
  return sheet;
}

function book(sheets: Sheet[]): Workbook {
  return { sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}

// ---------------------------------------------------------------- the pages

test('print preview: the sheets of a job are those printing takes: the shown one, the visible ones, a list', () => {
  const workbook = book([makeSheet('A', 3), makeSheet('B', 3, undefined, 'hidden'), makeSheet('C', 3), makeSheet('D', 3, undefined, 'veryHidden')]);
  assert.deepEqual(jobSheets(workbook, undefined, 2), [2]);
  assert.deepEqual(jobSheets(workbook, 'active', 2), [2]);
  assert.deepEqual(jobSheets(workbook, 'all', 2), [0, 2]);
  // The preview names the sheet that was shown when it opened: the same sheet.
  assert.deepEqual(jobSheets(workbook, [2], 0), jobSheets(workbook, 'active', 2));
  // A list keeps only the sheets the workbook still has.
  assert.deepEqual(jobSheets(workbook, [0, 7, 2], 0), [0, 2]);
});

test('print preview: the page sizes it shows are those of the plan printing makes, in order', () => {
  const workbook = book([makeSheet('A', 120), makeSheet('B', 30, { page: { orientation: 'landscape' } } as SheetPageSetup), makeSheet('C', 0)]);
  const plan = planSheetPages(workbook, jobSheets(workbook, 'all', 0), { defaultPaper: 9 });
  const sizes = plan.pages.map((entry) => entry.layout.paper);
  // Sheet C holds nothing: it has no pages, as when it is printed.
  assert.deepEqual(plan.sheets, [0, 1]);
  assert.ok(sizes.length >= 3);
  assert.equal(sizes[0].width < sizes[0].height, true);
  assert.equal(sizes[sizes.length - 1].width > sizes[sizes.length - 1].height, true);
  const layout = layoutPreview(sizes, 1, 1200);
  assert.equal(layout.slots.length, plan.pages.length);
  for (let i = 0; i < sizes.length; i++) {
    assert.equal(layout.slots[i].width, Math.round(sizes[i].width));
    assert.equal(layout.slots[i].height, Math.round(sizes[i].height));
  }
});

test('print preview: nothing to print and too many pages are what printing says', () => {
  const empty = book([makeSheet('A', 0)]);
  assert.throws(() => planSheetPages(empty, jobSheets(empty, 'active', 0)), (error: unknown) => error instanceof DocumentPrintError && error.code === 'empty');
  const long = book([makeSheet('A', 50)]);
  assert.throws(() => planSheetPages(long, [0], { maxPages: 1 }), (error: unknown) => error instanceof DocumentPrintError && error.code === 'too-long');
});

test('print preview: 2,000 pages are laid out at once, quickly', () => {
  // About 48 rows to an A4 page: just under the limit.
  const workbook = book([makeSheet('A', 95000)]);
  const started = Date.now();
  const plan = planSheetPages(workbook, [0], { defaultPaper: 9 });
  const layout = layoutPreview(
    plan.pages.map((entry) => entry.layout.paper),
    0.5,
    1280,
  );
  // A guard against a hang, not a measure of speed (about 10 ms when the machine is idle).
  assert.ok(Date.now() - started < 5000);
  assert.ok(plan.pages.length > 1900 && plan.pages.length <= MAX_SHEET_PAGES, String(plan.pages.length));
  assert.equal(layout.slots.length, plan.pages.length);
});

// --------------------------------------------------------------- the layout

test('print preview: at a whole page each page is a row of its own, centred, with the gap around', () => {
  const layout = layoutPreview([A4, A4, A4], 0.5, 800);
  assert.equal(layout.rows.length, 3);
  const w = Math.round(794 * 0.5);
  const h = Math.round(1123 * 0.5);
  // Two A4 pages at half size (397 + 16 + 397) do not fit in 800 - 32.
  for (let i = 0; i < 3; i++) {
    assert.deepEqual(layout.rows[i], { top: PREVIEW_GAP + i * (h + PREVIEW_GAP), bottom: PREVIEW_GAP + i * (h + PREVIEW_GAP) + h, first: i, last: i });
    assert.equal(layout.slots[i].width, w);
    assert.equal(layout.slots[i].x, Math.round(PREVIEW_GAP + (800 - 2 * PREVIEW_GAP - w) / 2));
  }
  assert.equal(layout.height, PREVIEW_GAP + 3 * (h + PREVIEW_GAP));
  assert.equal(layout.width, 800);
});

test('print preview: a small zoom puts pages side by side, as thumbnails', () => {
  const pages = new Array(10).fill(A4);
  const layout = layoutPreview(pages, 0.25, 1000);
  // 199 px each: four fit in 1000 - 32 with three gaps (4 * 199 + 3 * 16 = 844), five do not (1059).
  assert.deepEqual(
    layout.rows.map((r) => [r.first, r.last]),
    [
      [0, 3],
      [4, 7],
      [8, 9],
    ],
  );
  // The pages of a row share its top; the row is as high as its highest page.
  assert.equal(layout.slots[4].y, layout.slots[7].y);
  assert.ok(layout.slots[5].x > layout.slots[4].x);
});

test('print preview: a landscape page in a row of portrait ones is centred in its height; a narrow view still has one page a row', () => {
  const layout = layoutPreview([A4, A4_LANDSCAPE], 0.25, 2000);
  assert.equal(layout.rows.length, 1);
  const row0 = layout.rows[0];
  assert.equal(row0.bottom - row0.top, Math.round(1123 * 0.25));
  assert.equal(layout.slots[1].y, row0.top + Math.round((Math.round(1123 * 0.25) - Math.round(794 * 0.25)) / 2));
  const narrow = layoutPreview([A4, A4], 4, 100);
  assert.equal(narrow.rows.length, 2);
  // Wider than the view: the content is as wide as the page.
  assert.equal(narrow.width, Math.round(794 * 4) + 2 * PREVIEW_GAP);
});

test('print preview: damaged sizes read as A4 and a zoom out of range is brought in', () => {
  const layout = layoutPreview([{ width: NaN, height: -1 }, { width: Infinity, height: 5 }], 1, 900);
  assert.equal(layout.slots[0].width, 794);
  assert.equal(layout.slots[1].height, 1123);
  assert.equal(clampZoom(0), 1);
  assert.equal(clampZoom(NaN), 1);
  assert.equal(clampZoom(0.01), MIN_PREVIEW_ZOOM);
  assert.equal(clampZoom(100), MAX_PREVIEW_ZOOM);
  assert.equal(layoutPreview([A4], 99, 900).zoom, MAX_PREVIEW_ZOOM);
  assert.deepEqual(layoutPreview([], 1, 900).rows, []);
});

// ----------------------------------------------------------- virtualisation

test('print preview: of 2,000 pages a scroll position shows only the few in its window', () => {
  const pages = new Array(2000).fill(A4);
  const layout = layoutPreview(pages, 0.75, 900);
  // 842 px a page, 858 with the gap.
  const step = Math.round(1123 * 0.75) + PREVIEW_GAP;
  assert.deepEqual(visiblePages(layout, 0, 800), { first: 0, last: 0 });
  assert.deepEqual(visiblePages(layout, 0, 1000), { first: 0, last: 1 });
  // In the middle: the page under the top and those that reach into the view.
  const top = PREVIEW_GAP + 1000 * step + 100;
  assert.deepEqual(visiblePages(layout, top, 700), { first: 1000, last: 1000 });
  assert.deepEqual(visiblePages(layout, top, 800), { first: 1000, last: 1001 });
  assert.deepEqual(visiblePages(layout, top, 700, 400), { first: 999, last: 1001 });
  // At the end.
  assert.deepEqual(visiblePages(layout, layout.height - 800, 800), { first: 1999, last: 1999 });
  // Beyond the end (a stale scroll position): the last row, not nothing.
  assert.deepEqual(visiblePages(layout, layout.height * 2, 800), { first: 1999, last: 1999 });
  assert.equal(visiblePages(layoutPreview([], 1, 900), 0, 800), undefined);
});

test('print preview: thumbnails in view are whole rows', () => {
  const layout = layoutPreview(new Array(100).fill(A4), 0.25, 1000);
  const range = visiblePages(layout, 0, 600);
  assert.ok(range);
  assert.equal(range.first, 0);
  // Rows of four: the window ends at the end of a row.
  assert.equal((range.last + 1) % 4, 0);
});

test('print preview: the current page is the one a third down the view; a row keeps the one chosen in it', () => {
  const layout = layoutPreview(new Array(20).fill(A4), 0.5, 800);
  assert.equal(pageAtScroll(layout, 0, 600, 0), 0);
  assert.equal(pageAtScroll(layout, scrollForPage(layout, 7), 600, 0), 7);
  assert.equal(scrollForPage(layout, 0), 0);
  assert.equal(scrollForPage(layout, 99), scrollForPage(layout, 19));
  const thumbs = layoutPreview(new Array(20).fill(A4), 0.25, 1000);
  assert.equal(pageAtScroll(thumbs, 0, 600, 2), 2);
  assert.equal(pageAtScroll(thumbs, scrollForPage(thumbs, 9), 600, 2), 8);
});

// --------------------------------------------------------------------- zoom

test('print preview: "fit" shows a whole page; the steps go up and down from anywhere', () => {
  assert.equal(fitZoom(A4, 794 + 2 * PREVIEW_GAP, 2000), 1);
  assert.equal(fitZoom(A4, 2000, 1123 / 2 + 2 * PREVIEW_GAP), 0.5);
  assert.equal(fitZoom(A4, 10, 10), MIN_PREVIEW_ZOOM);
  assert.equal(stepZoom(1, 1), 1.1);
  assert.equal(stepZoom(1, -1), 0.9);
  assert.equal(stepZoom(0.8, 1), 0.9);
  assert.equal(stepZoom(0.8, -1), 0.75);
  assert.equal(stepZoom(MAX_PREVIEW_ZOOM, 1), MAX_PREVIEW_ZOOM);
  assert.equal(stepZoom(MIN_PREVIEW_ZOOM, -1), MIN_PREVIEW_ZOOM);
  for (let i = 1; i < PREVIEW_ZOOM_STEPS.length; i++) assert.ok(PREVIEW_ZOOM_STEPS[i] > PREVIEW_ZOOM_STEPS[i - 1]);
});

// ------------------------------------------------------------------- memory

test('print preview: pages are painted with the pixels of the screen, less when the budget is passed', () => {
  const pages = [A4, A4, A4];
  const layout = layoutPreview(pages, 1, 900);
  const budget = previewBudget(false);
  assert.deepEqual(previewPixelWidths(layout.slots, pages, [0, 1], 2, budget), [794 * 2, 794 * 2]);
  // A ratio that cannot be reads as 1.
  assert.deepEqual(previewPixelWidths(layout.slots, pages, [0], NaN, budget), [794]);
  // Too many pixels together: all are made smaller alike, and together stay within the budget.
  const tight = { pagePixels: 16e6, totalPixels: 1e6, canvases: 64 };
  const widths = previewPixelWidths(layout.slots, pages, [0, 1, 2], 2, tight);
  assert.equal(widths[0], widths[2]);
  let total = 0;
  for (const width of widths) total += width * ((width * 1123) / 794);
  assert.ok(total <= 1e6 + 3 * 1123, String(total));
  // One page alone above the limit of a canvas.
  const big = layoutPreview([A4], 4, 900);
  const [one] = previewPixelWidths(big.slots, [A4], [0], 3, { pagePixels: 4e6, totalPixels: 1e9, canvases: 64 });
  assert.ok(one * ((one * 1123) / 794) <= 4e6);
  assert.ok(previewBudget(true).totalPixels < previewBudget(false).totalPixels);
});

test('print preview: the canvas pool holds canvases only for pages in the window and lets go of their pixels', () => {
  let made = 0;
  const pool = new CanvasPool(() => {
    made++;
    return { width: 300, height: 150 };
  }, 2);
  const a = pool.take(0);
  a.width = 100;
  a.height = 200;
  assert.equal(pool.take(0), a);
  pool.take(1).width = 10;
  pool.take(2);
  pool.take(3);
  assert.equal(pool.size, 4);
  assert.equal(made, 4);
  const gone = pool.keepOnly((page) => page >= 2);
  assert.deepEqual(
    gone.map((entry) => entry.page),
    [0, 1],
  );
  assert.equal(a.width, 0);
  assert.equal(a.height, 0);
  assert.deepEqual(pool.pages(), [2, 3]);
  // A page that comes into the window takes a canvas that was given back.
  pool.take(10);
  assert.equal(made, 4);
  // No more than `spare` are kept: the rest are left to be collected.
  pool.keepOnly(() => false);
  pool.take(20);
  pool.take(21);
  pool.take(22);
  assert.equal(made, 5);
  pool.clear();
  assert.equal(pool.size, 0);
  assert.equal(pool.pixels, 0);
});

test('print preview: scrolling through 2,000 pages never holds more canvases than the window has pages', () => {
  const pages = new Array(2000).fill(A4);
  const layout = layoutPreview(pages, 0.25, 1280);
  const pool = new CanvasPool(() => ({ width: 0, height: 0 }));
  const budget = previewBudget(false);
  let most = 0;
  let mostPixels = 0;
  for (let top = 0; top < layout.height; top += 731) {
    const range = visiblePages(layout, top, 900, 450);
    assert.ok(range);
    pool.keepOnly((page) => page >= range.first && page <= range.last);
    const indexes: number[] = [];
    for (let page = range.first; page <= range.last; page++) indexes.push(page);
    const widths = previewPixelWidths(layout.slots, pages, indexes, 2, budget);
    indexes.forEach((page, i) => {
      const canvas = pool.take(page);
      canvas.width = widths[i];
      canvas.height = Math.round((widths[i] * 1123) / 794);
    });
    most = Math.max(most, pool.size);
    mostPixels = Math.max(mostPixels, pool.pixels);
  }
  assert.ok(most <= budget.canvases, String(most));
  assert.ok(most < 60, String(most));
  assert.ok(mostPixels <= budget.totalPixels * 1.01, String(mostPixels));
});

// ------------------------------------------------------------------- states

test('print preview: the states it goes through', () => {
  const run = (events: PreviewEvent[]): PreviewState[] => {
    const out: PreviewState[] = [];
    let state: PreviewState = 'closed';
    for (const event of events) out.push((state = previewTransition(state, event)));
    return out;
  };
  assert.deepEqual(run(['open', 'planned', 'changed', 'planned', 'close']), ['planning', 'ready', 'planning', 'ready', 'closed']);
  assert.deepEqual(run(['open', 'nothing', 'changed', 'error', 'changed', 'planned']), ['planning', 'empty', 'planning', 'failed', 'planning', 'ready']);
  // A closed preview hears nothing but 'open'; an open one does not open twice.
  assert.deepEqual(run(['planned', 'changed', 'open', 'open']), ['closed', 'closed', 'planning', 'planning']);
  // What comes after the pages are there does not change them until they are laid out again.
  assert.deepEqual(run(['open', 'planned', 'nothing', 'error']), ['planning', 'ready', 'ready', 'ready']);
  assert.deepEqual(run(['open', 'close', 'planned']), ['planning', 'closed', 'closed']);
});

// ------------------------------------------------------------------ strings

test('print preview: its words in both languages', () => {
  assert.equal(STRINGS_EN.preview.page(3, 12), 'Page 3 of 12');
  assert.equal(STRINGS_TR.preview.page(3, 12), 'Sayfa 3 / 12');
  assert.equal(STRINGS_EN.preview.counter(3, 12), '3 / 12');
  assert.equal(STRINGS_TR.preview.percent(75), '%75');
  assert.equal(STRINGS_EN.preview.percent(75), '75%');
  for (const strings of [STRINGS_EN.preview, STRINGS_TR.preview]) {
    for (const [key, value] of Object.entries(strings)) if (typeof value === 'string') assert.ok(value.length > 0, key);
  }
});

test('print preview: "&F" is the same file name for printing, the preview and the PDF: the one given, else that of the view', () => {
  // print() without a name used to leave "&F" empty while the preview and the PDF showed the name of the view.
  assert.equal(jobFileName(undefined, 'Kitap.xlsx'), 'Kitap.xlsx');
  assert.equal(jobFileName('Other.xlsx', 'Kitap.xlsx'), 'Other.xlsx');
  // A name given as empty is a name: "&F" is empty on purpose.
  assert.equal(jobFileName('', 'Kitap.xlsx'), '');
  assert.equal(jobFileName(undefined, undefined), undefined);
});

// ------------------------------------------------------------ review fixes

test('print preview: "active" is the sheet it was opened on, followed through moves and deletions; Print gets that list', () => {
  const a = makeSheet('A', 3);
  const b = makeSheet('B', 3);
  const c = makeSheet('C', 3);
  const workbook = book([a, b, c]);
  // The sheet shown now (another one, after the page setup of a page of B) does not count: A does.
  assert.deepEqual(previewSheets(workbook, 'active', a, 1), [0]);
  // A moved to the end, and the list follows it.
  workbook.sheets = [b, c, a];
  assert.deepEqual(previewSheets(workbook, 'active', a, 0), [2]);
  // A was deleted: the sheet that is shown.
  workbook.sheets = [b, c];
  assert.deepEqual(previewSheets(workbook, 'active', a, 1), [1]);
  assert.deepEqual(previewSheets(workbook, 'all', a, 1), [0, 1]);
});

test('print preview: a page painted while its pictures are still coming is painted again for 30 s of its layout at most, retries or not', () => {
  const clock = new SettleClock(30000);
  assert.equal(clock.again(1, 0), true);
  // A retry starts nothing over: the time runs from the first painting of the layout.
  assert.equal(clock.again(1, 20000), true);
  assert.equal(clock.again(1, 30001), false);
  assert.equal(clock.again(1, 90000), false);
  // A new layout (a zoom, a change of the workbook) has its own time.
  assert.equal(clock.again(2, 90000), true);
});

test('print preview: pictures beyond the budget are not decoded over and over: one pass a turn, nothing painted anew because one came', async () => {
  // Twelve large pictures that do not fit the budget together, painted the way the preview paints: a pass per turn,
  // pages that are not settled painted again every 120 ms, no new layout when a picture comes.
  const png = (): Uint8Array => {
    const bytes = new Uint8Array(33);
    bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
    const view = new DataView(bytes.buffer);
    view.setUint32(16, 4000);
    view.setUint32(20, 3000);
    bytes[24] = 8;
    bytes[25] = 6;
    return bytes;
  };
  let decodes = 0;
  let loads = 0;
  const store = new ImageStore({ readPart: async () => png() }, 64 * 1024 * 1024, () => loads++);
  (store as unknown as { decode: unknown }).decode = async (_bytes: unknown, _size: unknown, target: { width: number; height: number }) => {
    decodes++;
    await new Promise((resolve) => setTimeout(resolve, 2));
    return { image: { source: {}, width: target.width, height: target.height }, close() {} };
  };
  const clock = new SettleClock(600);
  const started = Date.now();
  let turns = 0;
  for (;;) {
    store.pass++;
    for (let page = 0; page < 12; page++) store.get(`xl/media/image${page}.png`, 4000, 3000);
    turns++;
    if (!store.busy || !clock.again(1, Date.now() - started)) break;
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
  const after = decodes;
  await new Promise((resolve) => setTimeout(resolve, 400));
  store.destroy();
  // It stops: nothing more is decoded once the time of the layout is over.
  assert.equal(decodes, after);
  assert.ok(turns <= 8, String(turns));
  assert.ok(loads >= 1);
});

test('print preview: every page in view gets its paper; only the canvases are limited, those in view first', () => {
  // A wide screen at the smallest zoom: more pages in view than canvases.
  const pages = new Array(2000).fill({ width: 559, height: 794 });
  const layout = layoutPreview(pages, 0.25, 2560);
  const window = previewWindow(layout, 0, 1320, 0, 24);
  const inView = visiblePages(layout, 0, 1320, 0)!;
  assert.ok(inView.last - inView.first + 1 > 24);
  // Paper for every page in view and around it.
  assert.ok(window.first <= inView.first && window.last >= inView.last);
  // A canvas for every page in view, even beyond the limit; none for the pages around it then.
  for (let page = inView.first; page <= inView.last; page++) assert.ok(window.canvases.includes(page), String(page));
  assert.equal(window.canvases.length, inView.last - inView.first + 1);
  // With room to spare, the pages around the view nearest the current one get canvases too, up to the limit.
  const small = previewWindow(layoutPreview(pages.slice(0, 50), 1, 900), 0, 600, 0, 4);
  assert.ok(small.canvases.length <= Math.max(4, small.visible));
  assert.ok(small.canvases.includes(0));
});

test('print preview: the largest zoom keeps the pages within what a browser can scroll', () => {
  assert.equal(maxPreviewZoom(new Array(10).fill(A4)), MAX_PREVIEW_ZOOM);
  const tall = new Array(2000).fill({ width: 794, height: 5000 });
  const most = maxPreviewZoom(tall);
  assert.ok(most < MAX_PREVIEW_ZOOM);
  assert.ok(layoutPreview(tall, most, 900).height <= 8_000_000 + 2000 * PREVIEW_GAP);
  // Never below the smallest zoom: past that the pages are what they are.
  assert.equal(maxPreviewZoom(new Array(2000).fill({ width: 794, height: 900000 })), MIN_PREVIEW_ZOOM);
});
