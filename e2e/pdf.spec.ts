// A PDF of a document and of a workbook: exportPdf() of the views, the
// button of the bars, the command of the WebView bridge. Every file that is
// made is taken out of the browser and read twice, by a reader of our own
// that follows the specification (test/pdf-fixtures) and, where the machine
// has them, by the tools of Poppler: the pages, their sizes, their pictures
// and the text that lies invisible over them.
//
// The first two groups are about pages that are pictures: their helper asks
// for `text: 'image'` unless a test says otherwise. (Their documents and
// workbooks are set in fonts of the device - Arial, Calibri and what stands
// in for it - whose files are not at hand, so their pages of text would be
// pictures by default too; a page without any text would not.) The last
// group has pages whose fonts have files - the font of the package and a
// registered one - and compares the two ways a page can be written.

import { readFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import type { Page } from '@playwright/test';
import { p } from '../test/doc-layout-fixtures';
import { buildDocx } from '../test/docx-fixtures';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { checkPdf, poppler, popplerOrSay, popplerPage, popplerWords, type CheckedPdf, type PopplerReport } from '../test/pdf-fixtures';
import { DOC_STYLES, docSection, fontStyles, picture } from './fixtures/docx';
import { testFont } from './fixtures/font';
import { expect, expectResponsive, test } from './helpers';

const MARGINS = '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>';
const A4 = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${MARGINS}</w:sectPr>`;
const A4_LANDSCAPE = `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>${MARGINS}</w:sectPr>`;
const LETTER = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr>`;
// Paper sizes in points, as the pages of the PDF state them.
const A4_PT = { width: 595.3, height: 841.9 };
const LETTER_PT = { width: 612, height: 792 };
const SHEET_A4_PT = { width: 595.28, height: 841.89 };

const SENTENCE = 'The quick brown fox jumps over the lazy dog and keeps running along the river bank. ';
const TURKISH = 'Çalışkan öğrenci ığdır’da şöyle üşüdü: İĞÜŞÖÇ';
const ASTRAL = 'Astral 😀 character';

/** `count` pages of A4, each with a heading and a few lines of text; the first one also has Turkish letters and a character outside the BMP. */
function pagesDocx(count: number, section = A4): Promise<Uint8Array> {
  let body = '';
  for (let i = 1; i <= count; i++) body += p(`Sheet ${i}`, i > 1 ? '<w:pageBreakBefore/>' : '') + (i === 1 ? p(TURKISH) + p(ASTRAL) : '') + p(SENTENCE + SENTENCE + SENTENCE);
  return buildDocx({ styles: DOC_STYLES, body: body + section });
}

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
const A4_SETUP = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9"/>';

/** 100 rows by 12 columns of numbers on A4: three pages down and two across. The first row starts with a word that fits its cell; a header and a footer. */
function sixPagesXlsx(more: { name: string; xml: string }[] = []): Promise<Uint8Array> {
  let data = '<sheetData>';
  for (let r = 1; r <= 100; r++) {
    data += `<row r="${r}">`;
    for (let c = 0; c < 12; c++) data += r === 1 && c === 0 ? str('A1', 'Ölçüm') : num(`${String.fromCharCode(65 + c)}${r}`, r * 100 + c);
    data += '</row>';
  }
  data += '</sheetData>' + A4_SETUP + '<headerFooter><oddHeader>&amp;LHEADER &amp;F&amp;RPage &amp;P of &amp;N</oddHeader><oddFooter>&amp;C&amp;A</oddFooter></headerFooter>';
  return buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(data) }, ...more] });
}

/** A sheet with one text in A1, in a column wide enough for it: what does not fit a cell at the edge of the sheet is cut off in print, and so in the file. */
const smallSheet = (name: string, text: string) => ({ name, xml: worksheetXml(`<cols><col min="1" max="1" width="30" customWidth="1"/></cols><sheetData><row r="1">${str('A1', text)}</row></sheetData>${A4_SETUP}`) });

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    c ^= bytes[i];
    for (let k = 0; k < 8; k++) c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
  }
  return (c ^ 0xffffffff) >>> 0;
}

/** A PNG that looks like a photograph to a compressor: soft ramps of colour with noise on them. Made by rule, the same every time. */
function photoPng(width: number, height: number): Uint8Array {
  const raw = new Uint8Array(height * (1 + width * 3));
  let seed = 12345;
  const noise = () => {
    seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
    return ((seed >>> 16) & 31) - 16;
  };
  const clamp = (value: number) => Math.max(0, Math.min(255, Math.round(value)));
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const at = y * (1 + width * 3) + 1 + x * 3;
      raw[at] = clamp(128 + 100 * Math.sin(x / 37) + noise());
      raw[at + 1] = clamp(128 + 100 * Math.cos(y / 53) + noise());
      raw[at + 2] = clamp(128 + 90 * Math.sin((x + y) / 71) + noise());
    }
  }
  const chunk = (type: string, data: Uint8Array): Uint8Array => {
    const out = new Uint8Array(12 + data.length);
    const view = new DataView(out.buffer);
    view.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
  };
  const header = new Uint8Array(13);
  new DataView(header.buffer).setUint32(0, width);
  new DataView(header.buffer).setUint32(4, height);
  header[8] = 8;
  header[9] = 2;
  return new Uint8Array(Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', deflateSync(raw)), chunk('IEND', new Uint8Array(0))]));
}

async function mountDoc(page: Page, bytes: Uint8Array, options: { toolbar?: boolean; locale?: string; fileName?: string } = {}): Promise<void> {
  await page.goto('/e2e/pages/doc.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountDoc(input, opts), [Array.from(bytes), options] as [number[], object]);
  await page.evaluate(() => ((window as any).__view = (window as any).__docView));
}

async function mountSheet(page: Page, bytes: Uint8Array, options: object = {}): Promise<void> {
  await page.goto('/e2e/pages/view.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountView(input, opts), [Array.from(bytes), { locale: 'en-US', ...options }] as [number[], object]);
  await expect(page.locator('.bym-canvas')).toBeVisible();
}

async function mountSheetEditor(page: Page, bytes: Uint8Array, options: object = {}): Promise<void> {
  await page.goto('/e2e/pages/editor.html');
  await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
  await page.evaluate(([input, opts]) => (window as any).__mountEditor(input, opts), [Array.from(bytes), { locale: 'en-US', ...options }] as [number[], object]);
  await expect(page.locator('#host .bym-canvas')).toBeVisible();
  await page.evaluate(() => ((window as any).__view = (window as any).__editor.view));
}

interface Outcome {
  /** The file as base64, when one was made. */
  data?: string;
  error?: { name: string; code?: string; message: string };
  progress: [number, number][];
}

/** Calls exportPdf() of window.__view (or of the object `holder` names) and brings the file, or the error, out of the page. Pages are pictures unless `options.text` says otherwise. */
function exportPdf(page: Page, options: object = {}, holder = '__view'): Promise<Outcome> {
  return page.evaluate(
    async ([opts, name]) => {
      const progress: [number, number][] = [];
      const target = (window as any)[name as string];
      try {
        const bytes: Uint8Array = await target.exportPdf({ text: 'image', ...(opts as object), onProgress: (done: number, total: number) => progress.push([done, total]) });
        // As base64, in steps: a file of megabytes is no argument list.
        let binary = '';
        for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
        return { data: btoa(binary), progress };
      } catch (error: any) {
        return { error: { name: String(error && error.name), code: error && error.code, message: String(error && error.message) }, progress };
      }
    },
    [options, holder] as [object, string],
  );
}

interface Read {
  bytes: Uint8Array;
  checked: CheckedPdf;
  /** What Poppler says of the file; undefined where it is not installed. */
  poppler: PopplerReport | undefined;
}

/** Reads a file with both readers. Poppler must have nothing to complain about. */
function read(data: string | undefined): Read {
  expect(data, 'a PDF was made').toBeTruthy();
  const bytes = new Uint8Array(Buffer.from(data as string, 'base64'));
  expect(Buffer.from(bytes.subarray(0, 5)).toString('latin1')).toBe('%PDF-');
  const checked = checkPdf(bytes);
  let report: PopplerReport | undefined;
  // Without Poppler the run says so once ("# poppler not found: ..."), and with BYM_REQUIRE_POPPLER=1 it fails here.
  if (popplerOrSay()) {
    report = poppler(bytes);
    expect(report.complaints, 'Poppler has nothing to say against the file').toBe('');
    expect(report.pages).toBe(checked.pages.length);
  }
  return { bytes, checked, poppler: report };
}

async function made(page: Page, options: object = {}, holder = '__view'): Promise<Read & { progress: [number, number][] }> {
  const outcome = await exportPdf(page, options, holder);
  expect(outcome.error, 'exportPdf() does not fail').toBeUndefined();
  return { ...read(outcome.data), progress: outcome.progress };
}

/** The words of a page as its text layer has them, in the order they are written. */
const words = (pdf: CheckedPdf, index: number): string => pdf.pages[index].text.replace(/\s+/g, ' ').trim();

const near = (actual: number, expected: number, tolerance: number) => Math.abs(actual - expected) <= tolerance;

test.describe('a document as a PDF', () => {
  test('exportPdf() gives a file with the pages of the document, each a picture with its text over it', async ({ page }) => {
    await mountDoc(page, await pagesDocx(3));
    await expect.poll(() => page.evaluate(() => (window as any).__view.layoutDone)).toBe(true);
    const count = await page.evaluate(() => (window as any).__view.pageCount);
    expect(count).toBe(3);
    const pdf = await made(page, { image: 'jpeg' });
    expect(pdf.checked.pages.length).toBe(count);
    for (const sheet of pdf.checked.pages) {
      expect([sheet.width, sheet.height]).toEqual([A4_PT.width, A4_PT.height]);
      // One picture that fills the page: A4 at 150 dots per inch.
      expect(sheet.images.length).toBe(1);
      expect([sheet.images[0].width, sheet.images[0].height, sheet.images[0].filter]).toEqual([1240, 1754, 'DCTDecode']);
      expect(sheet.images[0].matrix).toEqual([A4_PT.width, 0, 0, A4_PT.height, 0, 0]);
      // The text is there and is not painted.
      expect(sheet.texts.length).toBeGreaterThan(0);
      expect(sheet.texts.every((run) => run.mode === 3)).toBe(true);
    }
    expect(words(pdf.checked, 0)).toContain('Sheet 1');
    expect(words(pdf.checked, 0)).toContain(TURKISH);
    expect(words(pdf.checked, 0)).toContain(ASTRAL);
    expect(words(pdf.checked, 1)).toContain('Sheet 2');
    expect(words(pdf.checked, 2)).toContain('The quick brown fox');
    // The heading stands where it is painted: half an inch from the left, near the top of the page.
    const heading = pdf.checked.pages[1].texts[0];
    expect(heading.text.indexOf('Sheet')).toBe(0);
    expect(near(heading.x, 36, 0.5), `the heading starts at the left margin (${heading.x})`).toBe(true);
    expect(heading.y > A4_PT.height - 80 && heading.y < A4_PT.height - 36, `the heading is at the top (${heading.y})`).toBe(true);
    // A run is as wide in the file as it is on the page: the first line of the sentence ends before the right margin.
    const line = pdf.checked.pages[1].texts.find((run) => run.text.indexOf('The quick') === 0)!;
    expect(line.x + line.advance).toBeLessThanOrEqual(A4_PT.width - 36 + 1);
    expect(line.advance).toBeGreaterThan(100);
    expect(pdf.progress).toEqual([[1, 3], [2, 3], [3, 3]]);
    expect(pdf.checked.info.Creator).toBe('BYM Office');
    if (pdf.poppler) {
      expect(pdf.poppler.sizes).toEqual([A4_PT, A4_PT, A4_PT]);
      expect(pdf.poppler.images).toEqual([1, 2, 3].map((number) => ({ page: number, width: 1240, height: 1754, color: 'rgb', encoding: 'jpeg' })));
      expect(pdf.poppler.fonts).toEqual([{ name: 'BYMGlyphless', type: 'CID TrueType', embedded: true, subset: false, unicode: true }]);
      // The lines of the first page come out in their order, with every letter.
      const lines = pdf.poppler.text[0].split('\n').map((entry) => entry.trim()).filter((entry) => entry !== '');
      expect(lines.slice(0, 3)).toEqual(['Sheet 1', TURKISH, ASTRAL]);
      expect(lines[3].indexOf('The quick brown fox jumps')).toBe(0);
      expect(pdf.poppler.text[2]).toContain('Sheet 3');
    }
    await expectResponsive(page);
  });

  test('the options decide the resolution, the kind of picture and what the file says of itself', async ({ page }) => {
    await mountDoc(page, await pagesDocx(1));
    const lossless = await made(page, { dpi: 72, image: 'lossless', title: 'Ölçüm raporu (1)', author: 'A \\ B', date: new Date(Date.UTC(2026, 9, 13, 12, 0, 0)) });
    const image = lossless.checked.pages[0].images[0];
    // A4 at 72 dots per inch is a pixel per point; black text on white paper is grey.
    expect([image.width, image.height, image.filter]).toEqual([595, 842, 'FlateDecode']);
    expect(image.samples!.length).toBe(595 * 842 * (image.colorSpace === 'DeviceGray' ? 1 : 3));
    // The paper is white and something dark is written on it.
    expect(image.samples![0]).toBe(255);
    expect(image.samples!.some((value) => value < 80)).toBe(true);
    expect(lossless.checked.info).toMatchObject({ Title: 'Ölçüm raporu (1)', Author: 'A \\ B', CreationDate: 'D:20261013120000Z', Producer: 'BYM Office' });
    if (lossless.poppler) expect(lossless.poppler.info.Title).toBe('Ölçüm raporu (1)');
    // A better JPEG is a larger file; a higher resolution has more pixels.
    const low = await made(page, { dpi: 100, image: 'jpeg', quality: 0.3 });
    const high = await made(page, { dpi: 100, image: 'jpeg', quality: 1 });
    expect(low.checked.pages[0].images[0].width).toBe(827);
    expect(high.bytes.length).toBeGreaterThan(low.bytes.length);
    const fine = await made(page, { dpi: 300 });
    // 300 dots per inch, unless the device paints a page of that size with fewer pixels.
    expect(fine.checked.pages[0].images[0].width).toBeGreaterThanOrEqual(1240);
    expect(fine.checked.pages[0].images[0].width).toBeLessThanOrEqual(2480);
    // Whatever the resolution, the text is the same and lies at the same place.
    expect(words(fine.checked, 0)).toBe(words(low.checked, 0));
    expect(near(fine.checked.pages[0].texts[0].x, low.checked.pages[0].texts[0].x, 0.01)).toBe(true);
  });

  test('unless a kind of picture is asked for, a page of text is stored as its pixels and a page with a photograph as a JPEG', async ({ page }) => {
    // The first page is text; the second one is filled by a photograph.
    const body = p('Sheet 1') + p(SENTENCE + SENTENCE + SENTENCE + TURKISH) + p(picture(700, 950), '<w:pageBreakBefore/>') + A4;
    await mountDoc(page, await buildDocx({ styles: DOC_STYLES, body, rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', photoPng(350, 475)]] }));
    await expect.poll(() => page.evaluate(() => (window as any).__view.layoutDone)).toBe(true);
    const pdf = await made(page);
    expect(pdf.checked.pages.length).toBe(2);
    const text = pdf.checked.pages[0].images[0];
    // Every pixel as it was painted; black on white is grey. Text is compressed as it is: predictors would make it larger.
    expect([text.width, text.height, text.filter, text.predictor, text.colorSpace]).toEqual([1240, 1754, 'FlateDecode', undefined, 'DeviceGray']);
    expect(text.samples![0]).toBe(255);
    expect(text.samples!.some((value) => value < 80)).toBe(true);
    const photo = pdf.checked.pages[1].images[0];
    expect([photo.width, photo.height, photo.filter, photo.colorSpace]).toEqual([1240, 1754, 'DCTDecode', 'DeviceRGB']);
    if (pdf.poppler) expect(pdf.poppler.images.map((image) => [image.page, image.color, image.encoding])).toEqual([[1, 'gray', 'image'], [2, 'rgb', 'jpeg']]);
    // Asked for, each kind is what every page gets.
    const lossless = await made(page, { image: 'lossless' });
    expect(lossless.checked.pages.map((sheet) => sheet.images[0].filter)).toEqual(['FlateDecode', 'FlateDecode']);
    // The photograph is stored with the predictors of PNG, and read back by Poppler (above) as by us.
    expect(lossless.checked.pages.map((sheet) => sheet.images[0].predictor)).toEqual([undefined, true]);
    const jpeg = await made(page, { image: 'jpeg' });
    expect(jpeg.checked.pages.map((sheet) => sheet.images[0].filter)).toEqual(['DCTDecode', 'DCTDecode']);
    // And the choice was the smaller file on both pages.
    expect(pdf.bytes.length).toBeLessThan(lossless.bytes.length);
    expect(pdf.bytes.length).toBeLessThan(jpeg.bytes.length);
  });

  test('ten pages make a file of a size that can be sent around, in every kind and resolution', async ({ page }, info) => {
    test.setTimeout(300_000);
    // Each page of the first document is full of text; each page of the second holds a photograph and a few lines.
    let texts = '';
    let photos = '';
    for (let i = 1; i <= 10; i++) {
      texts += p(`Sheet ${i}`, i > 1 ? '<w:pageBreakBefore/>' : '');
      for (let k = 0; k < 13; k++) texts += p(SENTENCE + SENTENCE + SENTENCE + TURKISH);
      photos += p(`Sheet ${i}`, i > 1 ? '<w:pageBreakBefore/>' : '') + p(picture(700, 600)) + p(SENTENCE + SENTENCE + SENTENCE + TURKISH);
    }
    const documents: [string, Uint8Array][] = [
      ['text', await buildDocx({ styles: DOC_STYLES, body: texts + A4 })],
      ['photo', await buildDocx({ styles: DOC_STYLES, body: photos + A4, rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', photoPng(700, 600)]] })],
    ];
    const sizes: { [kind: string]: number } = {};
    for (const [name, bytes] of documents) {
      await mountDoc(page, bytes);
      await expect.poll(() => page.evaluate(() => (window as any).__view.layoutDone)).toBe(true);
      for (const dpi of [150, 300]) {
        for (const image of ['auto', 'jpeg', 'lossless']) {
          const kind = `${name} ${image} ${dpi}`;
          const outcome = await exportPdf(page, { dpi, image });
          expect(outcome.error, kind).toBeUndefined();
          const file = new Uint8Array(Buffer.from(outcome.data as string, 'base64'));
          const checked = checkPdf(file);
          expect(checked.pages.length, kind).toBe(10);
          sizes[kind] = file.length;
          // Guards against a file that is out of all proportion, not measurements: eight megabytes a page at most (noise, kept pixel by pixel, is as bad as it gets).
          expect(file.length, kind).toBeLessThan(80 * 1024 * 1024);
          expect(file.length, kind).toBeGreaterThan(20000);
          if (image === 'lossless' && dpi === 150) {
            // What the predictors are worth: the pictures as they are in the file, and the same samples compressed here without them.
            let stored = 0;
            let plain = 0;
            for (const sheet of checked.pages) {
              stored += sheet.images[0].data.length;
              plain += deflateSync(sheet.images[0].samples as Uint8Array).length;
            }
            sizes[`${name} lossless 150: pictures as stored (${checked.pages[0].images[0].predictor ? 'with' : 'without'} predictors)`] = stored;
            sizes[`${name} lossless 150: the same compressed without predictors`] = plain;
          }
        }
      }
      // Choosing page by page is never much worse than the better of the two kinds.
      for (const dpi of [150, 300]) expect(sizes[`${name} auto ${dpi}`], `${name} ${dpi}`).toBeLessThanOrEqual(Math.min(sizes[`${name} jpeg ${dpi}`], sizes[`${name} lossless ${dpi}`]) * 1.1);
    }
    info.annotations.push({ type: 'sizes', description: JSON.stringify(sizes) });
    console.log(`[pdf sizes, ${info.project.name}] ${JSON.stringify(sizes)}`);
  });

  test('pages of different sizes keep them, and an empty page holds no picture', async ({ page }) => {
    const body = p('Portrait') + p('', A4) + p('Landscape') + p('', A4_LANDSCAPE) + p('', '<w:pageBreakBefore/>') + LETTER;
    await mountDoc(page, await buildDocx({ styles: DOC_STYLES, body }));
    const pdf = await made(page, { dpi: 72 });
    expect(pdf.checked.pages.map((sheet) => [sheet.width, sheet.height])).toEqual([
      [A4_PT.width, A4_PT.height],
      [A4_PT.height, A4_PT.width],
      [LETTER_PT.width, LETTER_PT.height],
    ]);
    expect(pdf.checked.pages.map((sheet) => sheet.images.length)).toEqual([1, 1, 0]);
    expect(pdf.checked.pages[2].content).toBe('');
    if (pdf.poppler) expect(pdf.poppler.sizes).toEqual([A4_PT, { width: A4_PT.height, height: A4_PT.width }, LETTER_PT]);
  });

  test('a file nobody wants any more is given up, and the next one is made', async ({ page }) => {
    await mountDoc(page, await pagesDocx(6));
    const outcome = await page.evaluate(async () => {
      const view = (window as any).__view;
      const controller = new AbortController();
      const seen: number[] = [];
      try {
        await view.exportPdf({
          dpi: 72,
          signal: controller.signal,
          onProgress: (done: number) => {
            seen.push(done);
            if (done === 2) controller.abort();
          },
        });
        return { seen, code: 'no error', busy: view.exportingPdf };
      } catch (error: any) {
        return { seen, code: error.code, name: error.name, busy: view.exportingPdf };
      }
    });
    expect(outcome).toEqual({ seen: [1, 2], code: 'cancelled', name: 'PdfExportError', busy: false });
    // A signal that is aborted already: nothing is started.
    const never = await page.evaluate(async () => {
      const controller = new AbortController();
      controller.abort();
      return (window as any).__view.exportPdf({ signal: controller.signal }).then(
        () => 'made',
        (error: any) => error.code,
      );
    });
    expect(never).toBe('cancelled');
    await expect(page.locator('.bym-doc-pdf')).toHaveAttribute('aria-busy', 'false');
    expect((await made(page, { dpi: 72 })).checked.pages.length).toBe(6);
    // Destroying the view ends a file that is being made.
    const destroyed = await page.evaluate(async () => {
      const view = (window as any).__view;
      const job = view.exportPdf({ dpi: 72, onProgress: (done: number) => done === 1 && view.destroy() });
      return job.then(
        () => 'made',
        (error: any) => error.code,
      );
    });
    expect(destroyed).toBe('cancelled');
    expect(await page.evaluate(() => (window as any).__view.exportPdf().then(() => 'made', (error: any) => error.code))).toBe('cancelled');
  });

  test('an option that cannot be is refused, and a page without a canvas says that no PDF can be made', async ({ page }) => {
    await mountDoc(page, await pagesDocx(1));
    for (const options of [{ dpi: 0 }, { dpi: 'high' }, { quality: 2 }, { image: 'png' }, { text: 'sharp' }]) {
      const outcome = await exportPdf(page, options);
      expect(outcome.error && outcome.error.name, JSON.stringify(options)).toBe('RangeError');
    }
    expect(await page.evaluate(() => (window as any).__notices.filter((notice: any) => notice.kind === 'error'))).toEqual([]);
    // Canvases that give no context: what a page without them looks like.
    await page.evaluate(() => {
      HTMLCanvasElement.prototype.getContext = () => null;
    });
    const outcome = await exportPdf(page);
    expect(outcome.error).toMatchObject({ name: 'PdfExportError', code: 'unavailable' });
    expect(await page.evaluate(() => (window as any).__notices.slice(-1))).toEqual([{ message: 'A PDF cannot be made here', kind: 'error' }]);
  });

  test('the button of the bar downloads the PDF under the name of the file', async ({ page }) => {
    await mountDoc(page, await pagesDocx(2), { locale: 'tr', fileName: 'Ölçüm raporu.docx' });
    const button = page.locator('.bym-doc-pdf');
    await expect(button).toBeVisible();
    await expect(button).toHaveAttribute('aria-label', 'PDF olarak indir');
    await expect(button).toHaveAttribute('aria-busy', 'false');
    // Beside the Print button.
    const order = await page.evaluate(() => Array.from(document.querySelectorAll('.bym-doc-bar button')).map((element) => element.className.replace('bym-doc-btn ', '')));
    expect(order.indexOf('bym-doc-pdf')).toBe(order.indexOf('bym-doc-print') + 1);
    // Inside the bar, on a phone too, and large enough to press.
    const box = await page.evaluate(() => {
      const bar = document.querySelector('.bym-doc-bar')!.getBoundingClientRect();
      const pdf = document.querySelector('.bym-doc-pdf')!.getBoundingClientRect();
      const print = document.querySelector('.bym-doc-print')!.getBoundingClientRect();
      return { inside: pdf.left >= bar.left && pdf.right <= bar.right && pdf.top >= bar.top && pdf.bottom <= bar.bottom, after: pdf.left >= print.right, width: pdf.width };
    });
    expect(box.inside).toBe(true);
    expect(box.after).toBe(true);
    expect(box.width).toBeGreaterThanOrEqual(32);
    const [download] = await Promise.all([page.waitForEvent('download', { timeout: 60_000 }), button.click()]);
    // Safari hands the name on with its letters taken apart ("O" and its dots); the name is the same.
    expect(download.suggestedFilename().normalize('NFC')).toBe('Ölçüm raporu.pdf');
    const path = await download.path();
    const pdf = read((await readFile(path)).toString('base64'));
    expect(pdf.checked.pages.length).toBe(2);
    expect(words(pdf.checked, 0)).toContain('Sheet 1');
    await expect(button).toHaveAttribute('aria-busy', 'false');
    expect(await page.evaluate(() => (window as any).__notices.map((notice: any) => notice.message))).toEqual(['PDF hazırlanıyor…', 'PDF hazır']);
    // No link is left in the page.
    expect(await page.evaluate(() => document.querySelectorAll('a[download]').length)).toBe(0);
    // Printing is as it was: the pages are prepared and print() of the window is called.
    const printed = await page.evaluate(async () => {
      let calls = 0;
      window.print = () => {
        calls++;
        window.dispatchEvent(new Event('afterprint'));
      };
      const result = await (window as any).__view.print({ dpi: 72 });
      return { calls, pages: result.pages, left: document.querySelectorAll('.bym-print').length };
    });
    expect(printed).toEqual({ calls: 1, pages: 2, left: 0 });
  });

  test('text that lies beyond the edge of the paper is not in the file', async ({ page }) => {
    // A paragraph that starts to the left of the paper: its first letters are not on the page.
    const body = p('Whole line') + p('KENARDISINDA-KALAN-BASLANGIC and the rest of the line', '<w:ind w:left="-2400"/>') + A4;
    await mountDoc(page, await buildDocx({ styles: DOC_STYLES, body }));
    const pdf = await made(page, { dpi: 72 });
    const runs = pdf.checked.pages[0].texts.filter((run) => run.text.trim() !== '');
    expect(runs[0].text).toBe('Whole line');
    const cut = runs[1];
    const whole = 'KENARDISINDA-KALAN-BASLANGIC and the rest of the line';
    // What is left is the end of the line, from the first letter that shows, and it starts at the edge.
    expect(whole.endsWith(cut.text), `"${cut.text}" is the end of the line`).toBe(true);
    expect(cut.text.length).toBeLessThan(whole.length - 3);
    expect(cut.text.length).toBeGreaterThan(20);
    expect(cut.text).toContain('and the rest of the line');
    expect(near(cut.x, 0, 0.01), `the run starts at the edge of the paper (${cut.x})`).toBe(true);
    expect(pdf.checked.pages[0].text).not.toContain('KENARDISINDA');
    if (pdf.poppler) expect(pdf.poppler.text[0]).not.toContain('KENARD');
  });

  test('two files are made one after the other, and printing goes on beside them', async ({ page }) => {
    await mountDoc(page, await pagesDocx(3));
    const outcome = await page.evaluate(async () => {
      const view = (window as any).__view;
      const log: string[] = [];
      let prints = 0;
      window.print = () => {
        prints++;
        window.dispatchEvent(new Event('afterprint'));
      };
      const head = (bytes: Uint8Array) => String.fromCharCode.apply(null, Array.from(bytes.subarray(0, 5)));
      const first = view.exportPdf({ dpi: 72, onProgress: (done: number) => log.push('a' + done) });
      const second = view.exportPdf({ dpi: 72, onProgress: (done: number) => log.push('b' + done) });
      const printed = view.print({ dpi: 72 });
      const busy = [view.exportingPdf, view.printing];
      const [a, b, result] = await Promise.all([first, second, printed]);
      return { log, heads: [head(a), head(b)], same: a.length === b.length, busy, after: [view.exportingPdf, view.printing], prints, pages: result.pages, left: document.querySelectorAll('.bym-print').length };
    });
    // The second file waits for the first: their pages are not painted in turns.
    expect(outcome.log).toEqual(['a1', 'a2', 'a3', 'b1', 'b2', 'b3']);
    expect(outcome.heads).toEqual(['%PDF-', '%PDF-']);
    expect(outcome.same).toBe(true);
    expect(outcome.busy).toEqual([true, true]);
    expect(outcome.after).toEqual([false, false]);
    expect([outcome.prints, outcome.pages, outcome.left]).toEqual([1, 3, 0]);
  });

  test('what was typed is in the file, and a document that is shown read-only has one too', async ({ page }) => {
    await page.goto('/e2e/pages/doc-edit.html');
    await expect(page.locator('#state')).toHaveText('ready');
    await page.evaluate(([input]) => (window as any).__mountEdit(input, {}), [Array.from(await pagesDocx(1))] as [number[]]);
    const keys = await page.evaluate(() => Object.keys((window as any).__mounted));
    expect(keys).toContain('exportPdf');
    // At the start of the document.
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('Yazılan metin ');
    });
    const edited = await made(page, { dpi: 72 }, '__mounted');
    expect(words(edited.checked, 0)).toContain('Yazılan metin Sheet 1');
    // An RTF file is shown read-only (it is not converted): its pages are a PDF all the same.
    const rtf = new TextEncoder().encode('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\f0\\fs24 Hello from RTF\\par Second paragraph\\par}');
    await page.evaluate(([input]) => (window as any).__mountEdit(input, {}), [Array.from(rtf)] as [number[]]);
    expect(await page.evaluate(() => (window as any).__view.readOnly)).toBe(true);
    const shown = await made(page, { dpi: 72 }, '__mounted');
    expect(shown.checked.pages.length).toBe(1);
    expect(words(shown.checked, 0)).toContain('Hello from RTF');
    expect(words(shown.checked, 0)).toContain('Second paragraph');
  });
});

test.describe('a workbook as a PDF', () => {
  test('exportPdf() gives the pages printing lays the cells out on, with their text', async ({ page }) => {
    await mountSheet(page, await sixPagesXlsx(), { fileName: 'book.xlsx' });
    const pdf = await made(page);
    expect(pdf.checked.pages.length).toBe(6);
    for (const sheet of pdf.checked.pages) {
      expect([sheet.width, sheet.height]).toEqual([SHEET_A4_PT.width, SHEET_A4_PT.height]);
      expect([sheet.images[0].width, sheet.images[0].height]).toEqual([1240, 1754]);
    }
    // The first page: the header, the text of A1, the numbers of the first rows in their order, the footer.
    const first = words(pdf.checked, 0);
    expect(first.indexOf('HEADER book.xlsx')).toBe(0);
    expect(first).toContain('Page 1 of 6');
    expect(first).toContain('Ölçüm 101');
    expect(first).toMatch(/101 102 103 104 105 106 107 108 109 .*200 201 202/);
    expect(first.trim().endsWith('Data')).toBe(true);
    // A1 lies in the picture where it lies on the paper: at the left margin (0.7 inch), below the top one (0.75 inch).
    const a1 = pdf.checked.pages[0].texts.find((run) => run.text === 'Ölçüm')!;
    expect(near(a1.x, 0.7 * 72 + 1.5, 2.5), `A1 starts at the left margin (${a1.x})`).toBe(true);
    expect(a1.y < SHEET_A4_PT.height - 54 && a1.y > SHEET_A4_PT.height - 54 - 15, `A1 is in the first row (${a1.y})`).toBe(true);
    expect(a1.advance).toBeGreaterThan(20);
    expect(a1.advance).toBeLessThan(48);
    // The last of the three pages down the first columns holds rows 97 to 100.
    expect(words(pdf.checked, 2)).toMatch(/9700 .*10009/);
    expect(pdf.progress.map((entry) => entry[0])).toEqual([1, 2, 3, 4, 5, 6]);
    if (pdf.poppler) {
      expect(pdf.poppler.sizes).toEqual(new Array(6).fill(SHEET_A4_PT));
      const lines = pdf.poppler.text[0].split('\n').map((entry) => entry.trim()).filter((entry) => entry !== '');
      expect(lines[0]).toBe('HEADER book.xlsx Page 1 of 6');
      expect(lines[1].indexOf('Ölçüm 101 102')).toBe(0);
      expect(lines[2].indexOf('200 201 202')).toBe(0);
      expect(lines[lines.length - 1]).toBe('Data');
    }
    await expectResponsive(page);
  });

  test('text of a cell that the next cell cuts off, or the margin of the page, is in the file as far as it shows', async ({ page }) => {
    const secret = 'GIZLI-UZUN-METIN-KIRPILAN-KISIM-SONU-GORUNMEZ-12345678901';
    const margin = 'KENARA-KADAR-GORUNEN-METIN-SAYFA-BOSLUGUNDA-KESILEN-SON-98765432109-VE-DAHASI-BURADA-BITER';
    // A1 is cut off by B1. J1 is the last column of the sheet: what does not fit it is cut off where the cells of the page end.
    const data = `<sheetData><row r="1">${str('A1', secret)}${str('B1', 'komsu')}${str('J1', margin)}</row><row r="2">${str('A2', 'ikinci')}</row></sheetData>${A4_SETUP}`;
    await mountSheet(page, await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(data) }] }));
    const pdf = await made(page);
    expect(pdf.checked.pages.length).toBe(1);
    const all = pdf.checked.pages[0].text;
    // The start of each text is there; what cannot be seen is not.
    expect(all).toContain('GIZLI-U');
    expect(all).not.toContain('GORUNMEZ');
    expect(all).not.toContain('12345678901');
    expect(all).toContain('KENARA');
    expect(all).not.toContain('BURADA-BITER');
    expect(all).not.toContain('98765432109');
    const runs = pdf.checked.pages[0].texts;
    const a1 = runs.find((run) => run.text.indexOf('GIZLI') === 0)!;
    const b1 = runs.find((run) => run.text === 'komsu')!;
    // A1 is 64 px (48 pt) wide: the run of its text ends where B1 begins, and the two are kept apart by a space.
    expect(a1.text.length).toBeLessThan(14);
    expect(a1.x + a1.advance).toBeLessThanOrEqual(b1.x + 0.01);
    expect(a1.advance).toBeLessThanOrEqual(48);
    expect(runs[runs.indexOf(a1) + 1].text).toBe(' ');
    // The text of J1 ends where the cells do, and before the right margin (0.7 inch).
    const j1 = runs.find((run) => run.text.indexOf('KENARA') === 0)!;
    expect(j1.x + j1.advance).toBeLessThanOrEqual(SHEET_A4_PT.width - 0.7 * 72 + 0.01);
    expect(j1.text.length).toBeGreaterThan(5);
    if (pdf.poppler) {
      const lines = pdf.poppler.text[0].split('\n').map((entry) => entry.trim()).filter((entry) => entry !== '');
      expect(lines[0]).toMatch(/^GIZLI-U\S* komsu KENARA\S*$/);
      expect(lines[1]).toBe('ikinci');
      expect(pdf.poppler.text[0]).not.toContain('12345678901');
      expect(pdf.poppler.text[0]).not.toContain('98765432109');
    }
  });

  test('cells with several lines are copied cell by cell, not line by line across the row', async ({ page }) => {
    const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    // One cell format: text wraps, and starts at the top of the cell.
    const styles =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs></styleSheet>';
    const wrapped = (ref: string, text: string) => `<c r="${ref}" s="1" t="inlineStr"><is><t>${text}</t></is></c>`;
    const data = `<sheetData><row r="1" ht="60" customHeight="1">${wrapped('A1', 'Gemi adı Deniz Yıldızı')}${wrapped('B1', 'Liman adı İzmir Ege')}</row><row r="2">${str('A2', 'alt')}${str('B2', 'satır')}</row></sheetData>${A4_SETUP}`;
    await mountSheet(page, await buildXlsx({ styles, sheets: [{ name: 'S', xml: worksheetXml(data) }] }));
    const pdf = await made(page, { dpi: 72 });
    const runs = pdf.checked.pages[0].texts.filter((run) => run.text.trim() !== '');
    // Both cells have more than one line.
    expect(new Set(runs.filter((run) => /Gemi|Deniz|Yıldızı/.test(run.text)).map((run) => Math.round(run.y))).size).toBeGreaterThan(1);
    expect(words(pdf.checked, 0)).toBe('Gemi adı Deniz Yıldızı Liman adı İzmir Ege alt satır');
    if (pdf.poppler) expect(pdf.poppler.text[0].replace(/\s+/g, ' ').trim()).toBe('Gemi adı Deniz Yıldızı Liman adı İzmir Ege alt satır');
  });

  test('"sheets" chooses the sheets, and a sheet with nothing on it gives no file', async ({ page }) => {
    await mountSheet(page, await sixPagesXlsx([smallSheet('Second', 'İkinci sayfa'), smallSheet('Third', 'Üçüncü sayfa'), { name: 'Empty', xml: worksheetXml('<sheetData/>') }]));
    const all = await made(page, { sheets: 'all', dpi: 72 });
    expect(all.checked.pages.length).toBe(8);
    expect(words(all.checked, 6)).toBe('İkinci sayfa');
    expect(words(all.checked, 7)).toBe('Üçüncü sayfa');
    const listed = await made(page, { sheets: [2, 1], dpi: 72 });
    expect(listed.checked.pages.map((_sheet, index) => words(listed.checked, index))).toEqual(['Üçüncü sayfa', 'İkinci sayfa']);
    for (const sheets of [[9], [-1], [1.5], [], 'some', ['1']]) {
      const outcome = await exportPdf(page, { sheets });
      expect(outcome.error && outcome.error.name, JSON.stringify(sheets)).toBe('RangeError');
    }
    const empty = await exportPdf(page, { sheets: [3] });
    expect(empty.error).toMatchObject({ name: 'PdfExportError', code: 'empty' });
    expect(await page.evaluate(() => (window as any).__notices.slice(-1))).toEqual([{ message: 'There is nothing to make a PDF of', kind: 'error' }]);
    // Printing is as it was: its own "sheets" still knows two words only.
    expect(await page.evaluate(() => (window as any).__view.print({ sheets: [1] }).then(() => 'printed', (error: any) => error.name))).toBe('RangeError');
  });

  test('what is being typed in a cell is applied first, and a read-only view makes a file too', async ({ page }) => {
    await mountSheet(page, await buildXlsx({ sheets: [smallSheet('S', 'Before')] }));
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.select({ top: 0, left: 1, bottom: 0, right: 1 });
      view.startEditing('Yeni');
    });
    const pdf = await made(page, { dpi: 72 });
    expect(words(pdf.checked, 0)).toBe('Before Yeni');
    expect(await page.evaluate(() => (window as any).__view.isEditing)).toBe(false);
    await mountSheet(page, await buildXlsx({ sheets: [smallSheet('S', 'Salt okunur')] }), { readOnly: true });
    expect(words((await made(page, { dpi: 72 })).checked, 0)).toBe('Salt okunur');
  });

  test('the button of the toolbar downloads the PDF, and says so while it is made', async ({ page }) => {
    await mountSheetEditor(page, await sixPagesXlsx(), { fileName: 'Kitap 1.xlsx' });
    const button = page.locator('#host .bym-tb-btn[data-command="pdf"]');
    await expect(button).toHaveAttribute('aria-label', 'Download as PDF');
    await expect(button).toHaveAttribute('aria-busy', 'false');
    const labels = await page.evaluate(() => Array.from(document.querySelectorAll('#host .bym-toolbar button')).map((element) => element.getAttribute('aria-label')));
    expect(labels.indexOf('Download as PDF')).toBe(labels.indexOf('Print') + 1);
    await page.evaluate(() => {
      const w = window as any;
      w.__status = [];
      const status = document.querySelector('#host .bym-status');
      if (status) new MutationObserver(() => w.__status.push(status.textContent)).observe(status, { childList: true, subtree: true, characterData: true });
    });
    await button.scrollIntoViewIfNeeded();
    const [download] = await Promise.all([page.waitForEvent('download', { timeout: 60_000 }), button.click()]);
    expect(download.suggestedFilename()).toBe('Kitap 1.pdf');
    const pdf = read((await readFile(await download.path())).toString('base64'));
    expect(pdf.checked.pages.length).toBe(6);
    expect(words(pdf.checked, 0)).toContain('HEADER Kitap 1.xlsx');
    await expect(button).toHaveAttribute('aria-busy', 'false');
    expect(await page.evaluate(() => (window as any).__events.filter((event: any) => event.type === 'notice').map((event: any) => event.message))).toEqual(['Making the PDF…', 'The PDF is ready']);
  });
});

test.describe('a PDF through the WebView bridge', () => {
  const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');

  async function startBridge(page: Page): Promise<void> {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: string) => (window as any).__messages.push(JSON.parse(json)) };
    });
    await page.goto('/dist/bymoffice.webview.html');
  }
  const send = (page: Page, message: object) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
  const answer = async (page: Page, id: string) => {
    let found: any;
    await expect
      .poll(async () => (found = await page.evaluate((wanted) => (window as any).__messages.find((message: any) => message.id === wanted), id)), { timeout: 60_000, message: `answer to ${id}` })
      .toBeTruthy();
    return found;
  };

  test('exportPdf is answered with the file, of a document and of a workbook', async ({ page }) => {
    await startBridge(page);
    await send(page, { type: 'exportPdf', id: 'none' });
    expect(await answer(page, 'none')).toMatchObject({ type: 'error', code: 'not-loaded' });
    await send(page, { type: 'load', id: 'l', name: 'Rapor.docx', data: base64(await pagesDocx(2)) });
    expect((await answer(page, 'l')).type).toBe('loaded');
    await send(page, { type: 'exportPdf', id: 'p', dpi: 72, title: 'Başlık' });
    // The messages after it do not wait for the file.
    await send(page, { type: 'setLocale', id: 's', locale: 'tr' });
    expect(await answer(page, 's')).toEqual({ type: 'done', id: 's', command: 'setLocale' });
    const reply = await answer(page, 'p');
    // The document is set in a font of the device: none of its pages is drawn.
    expect({ type: reply.type, name: reply.name, pages: reply.pages, vector: reply.vector }).toEqual({ type: 'pdf', name: 'Rapor.pdf', pages: 2, vector: 0 });
    const pdf = read(reply.data);
    expect(pdf.checked.pages.length).toBe(2);
    expect(pdf.checked.info.Title).toBe('Başlık');
    expect(words(pdf.checked, 1)).toContain('Sheet 2');
    for (const [id, wrong] of [['b1', { dpi: 'high' }], ['b2', { image: 'png' }], ['b3', { quality: 0 }], ['b4', { sheets: 'some' }], ['b5', { title: 5 }], ['b6', { text: 'sharp' }]] as [string, object][]) {
      // (and a way of writing pages that is none)
      await send(page, { type: 'exportPdf', id, ...wrong });
      expect(await answer(page, id)).toMatchObject({ type: 'error', id, code: 'bad-message' });
    }
    // Only text and drawings are asked for, and the document is in a font of the device.
    await send(page, { type: 'exportPdf', id: 'v', text: 'vector' });
    const refused = await answer(page, 'v');
    expect({ type: refused.type, code: refused.code }).toEqual({ type: 'error', code: 'pdf-unsupported' });
    expect(refused.message).toMatch(/^Page 1 cannot be written as text and drawings: /);
    // Closing the file while its pages are painted.
    await send(page, { type: 'exportPdf', id: 'p2' });
    await send(page, { type: 'close', id: 'c' });
    expect(await answer(page, 'p2')).toMatchObject({ type: 'error', id: 'p2', code: 'pdf-failed' });
    // A workbook: the sheet that is shown, or all of them; the name of the file is what "&F" stands for.
    await send(page, { type: 'load', id: 'x', name: 'Kitap.xlsx', data: base64(await sixPagesXlsx([smallSheet('Second', 'İkinci sayfa')])) });
    await answer(page, 'x');
    await send(page, { type: 'exportPdf', id: 'w', dpi: 72, image: 'lossless', sheets: 'all' });
    const workbook = await answer(page, 'w');
    expect({ type: workbook.type, name: workbook.name, pages: workbook.pages }).toEqual({ type: 'pdf', name: 'Kitap.pdf', pages: 7 });
    const sheets = read(workbook.data);
    expect(words(sheets.checked, 0).indexOf('HEADER Kitap.xlsx')).toBe(0);
    expect(words(sheets.checked, 6)).toBe('İkinci sayfa');
    expect(sheets.checked.pages[0].images[0].filter).toBe('FlateDecode');
  });
});

// ---------------------------------------------------------------------------
// Pages written as text and drawings

/** A family no device has: the font of the package is told to stand in for it. */
const ABSENT = 'Bym Absent Bookman';
/** A family for the font the tests make (every character a box one em wide). */
const BOXES = 'Bym Pdf Boxes';
const inFont = (text: string, family: string, more = '') => `<w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>${more}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (runs: string, properties = '') => `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${runs}</w:p>`;

/** Words with letters fonts join (fi, fl, ff, ffi, ffl), and one with letters they do not. */
const JOINED = 'fiyat ofis raflar office waffle difficult fjord';

/** `count` pages set in `family`, 12 pt: a heading in bold, a line with Turkish letters, with `joined` a line of words whose letters fonts join, a line in italics, a shaded paragraph with a border and a few lines of text. */
function fontDocx(count: number, family: string, extraOnPage?: (page: number) => string, joined = false): Promise<Uint8Array> {
  let body = '';
  for (let i = 1; i <= count; i++) {
    body += para(inFont(`Bölüm ${i}`, family, '<w:b/>'), i > 1 ? '<w:pageBreakBefore/>' : '');
    body += para(inFont(TURKISH, family));
    if (joined) body += para(inFont(JOINED, family) + inFont(' ' + JOINED, family, '<w:b/>'));
    body += para(inFont('Eğik yazılmış bir satır, ', family, '<w:i/>') + inFont('kalın eğik', family, '<w:b/><w:i/>') + inFont(' ve altı çizili', family, '<w:u w:val="single"/>'));
    body += para(inFont('Gölgeli ve kenarlıklı paragraf', family), '<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/><w:bottom w:val="single" w:sz="8" w:space="1" w:color="0000FF"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/>');
    body += para(inFont(SENTENCE + SENTENCE + SENTENCE, family));
    if (extraOnPage) body += extraOnPage(i);
  }
  return buildDocx({ styles: fontStyles(family, 24), body: body + A4 });
}

/** Opens the page of documents with the font of the package standing in for the family no device has. */
async function openDocPage(page: Page): Promise<void> {
  await page.goto('/e2e/pages/doc.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate((names) => (window as any).BymOffice.fonts.configure({ as: names }), ['Bookman Old Style', ABSENT]);
}

async function showDoc(page: Page, bytes: Uint8Array, options: object = {}): Promise<void> {
  await page.evaluate(([input, opts]) => (window as any).__mountDoc(input, opts), [Array.from(bytes), options] as [number[], object]);
  await page.evaluate(() => ((window as any).__view = (window as any).__docView));
}

/** Waits until the font of the package is on the page and the document was laid out with it. */
async function withPackageFont(page: Page): Promise<void> {
  await expect.poll(() => page.evaluate(() => (window as any).BymOffice.fonts.embedded().status as string), { timeout: 20_000 }).toBe('loaded');
  await expect.poll(() => page.evaluate(() => (window as any).__view.layoutDone), { timeout: 60_000 }).toBe(true);
}

/** exportPdf() with what onPage is told. */
function exportWithPages(page: Page, options: object = {}): Promise<Outcome & { pages: { page: number; mode: string; reason?: string }[] }> {
  return page.evaluate(async (opts) => {
    const pages: { page: number; mode: string; reason?: string }[] = [];
    try {
      const bytes: Uint8Array = await (window as any).__view.exportPdf({ ...(opts as object), onPage: (info: any) => pages.push({ page: info.page, mode: info.mode, reason: info.reason }) });
      let binary = '';
      for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
      return { data: btoa(binary), progress: [], pages };
    } catch (error: any) {
      return { error: { name: String(error && error.name), code: error && error.code, message: String(error && error.message), page: error && error.page, reason: error && error.reason } as any, progress: [], pages };
    }
  }, options);
}

interface Raster {
  width: number;
  height: number;
  data: Uint8Array;
}

/** A band of rows of a picture with something on it (a line of text, a border): its first and last row, and its ink. */
interface InkBand {
  top: number;
  bottom: number;
  ink: number;
}

/**
 * The bands of rows of a picture that hold ink: a line of text is one (with
 * what is drawn in its rows: an underline, a shading), a border between two
 * lines another. The ink of a row is the darkness of its pixels summed, one
 * for a black pixel; a row counts when it holds a twenty-fifth of what the
 * fullest row holds - more than the lines that run down a table leave in a
 * row. (Poppler paints a line thinner than a pixel a whole pixel wide and
 * in its full colour, where the picture of the page has it faint: such a
 * line must not decide where the bands are.)
 */
function inkBands(picture: Raster): { bands: InkBand[]; total: number } {
  const rows = new Array<number>(picture.height).fill(0);
  for (let y = 0; y < picture.height; y++) {
    for (let x = 0; x < picture.width; x++) {
      const at = (y * picture.width + x) * 3;
      const dark = 1 - (picture.data[at] * 0.299 + picture.data[at + 1] * 0.587 + picture.data[at + 2] * 0.114) / 255;
      if (dark > 0.02) rows[y] += dark;
    }
  }
  const least = Math.max(3, Math.max(...rows) / 25);
  const bands: InkBand[] = [];
  let total = 0;
  for (let y = 0; y < rows.length; y++) {
    total += rows[y];
    if (rows[y] < least) continue;
    const last = bands[bands.length - 1];
    // Two rows without ink do not end a band: the dots of letters belong to their line.
    if (last && y - last.bottom <= 3) {
      last.bottom = y;
      last.ink += rows[y];
    } else bands.push({ top: y, bottom: y, ink: rows[y] });
  }
  return { bands, total };
}

/**
 * Where two pictures of one page differ in what is on them, line by line:
 * the same bands of rows hold ink in both - as many, each starting and
 * ending within four pixels (the dots over letters are faint rows that
 * count in one picture and not in the other) - and each band holds the same share of the ink
 * of its page within a quarter (two painters smooth the edges of letters
 * differently, and one may paint all text a little bolder: that changes
 * every line alike). A line of text that is missing or twice there, one that
 * lies more than four pixels higher or lower, one set in another size or cut
 * short is a difference. Tried on pages made in Node: a page against its own
 * picture has none; with a line of forty missing, six pixels lower, or cut
 * to half its words it has. A line moved sideways is not seen here: the
 * boxes of the words are compared for that.
 */
function inkDifferences(a: Raster, b: Raster): string[] {
  expect([a.width, a.height]).toEqual([b.width, b.height]);
  const first = inkBands(a);
  const second = inkBands(b);
  if (first.bands.length !== second.bands.length) return [`${first.bands.length} bands of rows with ink against ${second.bands.length}`];
  const out: string[] = [];
  for (let i = 0; i < first.bands.length; i++) {
    const one = first.bands[i];
    const other = second.bands[i];
    if (Math.abs(one.top - other.top) > 4 || Math.abs(one.bottom - other.bottom) > 4) out.push(`band ${i + 1}: rows ${one.top} to ${one.bottom} against ${other.top} to ${other.bottom}`);
    const share = one.ink / first.total;
    const otherShare = other.ink / second.total;
    // A band of a row or two is a thin line: where it lies is compared, its ink is spread over its rows as the painter likes.
    if (one.bottom - one.top >= 3 && Math.abs(share - otherShare) > Math.max(share, otherShare) * 0.25 + 0.0005) out.push(`band ${i + 1} (rows ${one.top} to ${one.bottom}): ${(share * 100).toFixed(2)} % of the ink against ${(otherShare * 100).toFixed(2)} %`);
  }
  return out.slice(0, 12);
}

/** The box around the pixels of a picture that are of a colour, and how many there are. */
function colourBox(picture: Raster, is: (r: number, g: number, b: number) => boolean): { left: number; top: number; right: number; bottom: number; count: number } {
  const box = { left: picture.width, top: picture.height, right: -1, bottom: -1, count: 0 };
  for (let y = 0; y < picture.height; y++) {
    for (let x = 0; x < picture.width; x++) {
      const at = (y * picture.width + x) * 3;
      if (!is(picture.data[at], picture.data[at + 1], picture.data[at + 2])) continue;
      box.count++;
      box.left = Math.min(box.left, x);
      box.right = Math.max(box.right, x);
      box.top = Math.min(box.top, y);
      box.bottom = Math.max(box.bottom, y);
    }
  }
  return box;
}

const RED = (r: number, g: number, b: number): boolean => r > 170 && g < 100 && b < 100;
const BLUE = (r: number, g: number, b: number): boolean => b > 170 && r < 100 && g < 100;
/** The shading of the bordered paragraph (#DDEEFF), as a picture keeps it. */
const SHADING = (r: number, g: number, b: number): boolean => r > 205 && r < 235 && g > 225 && g < 250 && b > 245;

/**
 * Thin and coloured things of the page (the two borders of the shaded
 * paragraph, and its shading) lie at the same place in both pictures: what a
 * comparison of ink by bands is too coarse for.
 */
function expectSameMarks(a: Raster, b: Raster, what: string): void {
  for (const [name, is] of [['the red border', RED], ['the blue border', BLUE], ['the shading', SHADING]] as [string, (r: number, g: number, b: number) => boolean][]) {
    const first = colourBox(a, is);
    const second = colourBox(b, is);
    expect(first.count, `${what}: ${name} is on the drawn page`).toBeGreaterThan(20);
    expect(second.count, `${what}: ${name} is on the picture of the page`).toBeGreaterThan(20);
    for (const edge of ['left', 'top', 'right', 'bottom'] as const) {
      expect(Math.abs(first[edge] - second[edge]), `${what}: the ${edge} edge of ${name} (${first[edge]} against ${second[edge]})`).toBeLessThanOrEqual(2);
    }
  }
}

interface WordBox {
  text: string;
  xMin: number;
  yMin: number;
  xMax: number;
  yMax: number;
}

/** The words of a page in their lines, from the top; within a line from the left. */
function linesOfWords(words: WordBox[]): WordBox[][] {
  const sorted = words.slice().sort((a, b) => (a.yMin + a.yMax) / 2 - (b.yMin + b.yMax) / 2);
  const lines: WordBox[][] = [];
  for (const word of sorted) {
    const line = lines[lines.length - 1];
    const middle = (word.yMin + word.yMax) / 2;
    if (line && Math.abs((line[0].yMin + line[0].yMax) / 2 - middle) < 5) line.push(word);
    else lines.push([word]);
  }
  return lines.map((line) => line.sort((a, b) => a.xMin - b.xMin));
}

/**
 * The words of the drawn page stand where the picture of the page has them:
 * the same words in the same lines, every line starting within a point and a
 * half and ending where the picture ends it, at the same height. The invisible text of a picture
 * gives every character of a run the same width, so the edges of the words
 * inside a run of text are those of the run only at its two ends; they are
 * held to the run loosely.
 */
function expectSameWords(drawn: WordBox[], pictured: WordBox[], what: string, cut = 0): void {
  const first = linesOfWords(drawn);
  const second = linesOfWords(pictured);
  expect(first.map((line) => line.map((word) => word.text).join(' ')), `${what}: the lines of the drawn page`).toEqual(second.map((line) => line.map((word) => word.text).join(' ')));
  for (let i = 0; i < first.length; i++) {
    const a = first[i];
    const b = second[i];
    const text = a.map((word) => word.text).join(' ');
    expect(Math.abs(a[0].xMin - b[0].xMin), `${what}: "${text}" starts at ${a[0].xMin}, in the picture at ${b[0].xMin}`).toBeLessThanOrEqual(1.5);
    // Where a line ends: a space at the end of a run takes its share of the run in the invisible text (a
    // character in sixty, say), and a letter a clip cuts through is whole in the drawn text where the invisible
    // text ends at the clip (`cut`: how wide a letter may be).
    const end = a[a.length - 1].xMax - b[b.length - 1].xMax;
    expect(end >= -Math.max(1.5, (a[a.length - 1].xMax - a[0].xMin) / 60) && end <= Math.max(1.5, (a[a.length - 1].xMax - a[0].xMin) / 60) + cut, `${what}: "${text}" ends at ${a[a.length - 1].xMax}, in the picture at ${b[b.length - 1].xMax}`).toBe(true);
    expect(Math.abs((a[0].yMin + a[0].yMax) / 2 - (b[0].yMin + b[0].yMax) / 2), `${what}: the height of "${text}"`).toBeLessThanOrEqual(4);
    // Inside a run the invisible text of the picture gives every character the same width: a word may stand a
    // sixteenth of its line away from where its letters are.
    const slack = Math.max(14, (a[a.length - 1].xMax - a[0].xMin) / 16);
    for (let k = 0; k < a.length; k++) expect(Math.abs(a[k].xMin - b[k].xMin), `${what}: "${a[k].text}" in "${text}" starts at ${a[k].xMin}, in the picture at ${b[k].xMin}`).toBeLessThanOrEqual(slack);
  }
}

test.describe('pages written as text and drawings', () => {
  test('a document in the font of the package is written as text in that font, and looks as its picture does', async ({ page }) => {
    await openDocPage(page);
    await showDoc(page, await fontDocx(2, ABSENT, undefined, true));
    await withPackageFont(page);
    const vector = await exportWithPages(page, { text: 'vector' });
    // Which ligatures the browser draws is found out by painting with kerning turned off. A browser whose canvas
    // cannot do that leaves a page with letters its font could join a picture, and says so.
    const canAsk = await page.evaluate(() => 'fontKerning' in (document.createElement('canvas').getContext('2d') as object));
    if (!canAsk) {
      expect(vector.error).toMatchObject({ name: 'PdfExportError', code: 'unsupported', page: 1 });
      expect((vector.error as any).reason).toMatch(/which ligatures of the font TeXGyreBonum-\w+ the browser draws/);
      const auto = await exportWithPages(page, {});
      expect(auto.pages.map((entry) => entry.mode)).toEqual(['image', 'image']);
      expect(words(read(auto.data).checked, 0)).toContain(JOINED + ' ' + JOINED);
      return;
    }
    expect(vector.error, 'every page can be drawn').toBeUndefined();
    expect(vector.pages).toEqual([
      { page: 1, mode: 'vector', reason: undefined },
      { page: 2, mode: 'vector', reason: undefined },
    ]);
    const drawn = read(vector.data);
    expect(drawn.checked.pages.length).toBe(2);
    for (const sheet of drawn.checked.pages) {
      expect([sheet.width, sheet.height]).toEqual([A4_PT.width, A4_PT.height]);
      expect(sheet.images.length, 'no picture of the page').toBe(0);
      expect(sheet.texts.length).toBeGreaterThan(4);
      expect(sheet.texts.every((run) => run.mode === 0), 'the text is painted').toBe(true);
      expect(sheet.shapes.length, 'the shading, the borders and the underline are drawings').toBeGreaterThan(3);
    }
    expect(words(drawn.checked, 0)).toContain(TURKISH);
    // Whatever ligatures the browser drew, the text is read back letter by letter.
    expect(words(drawn.checked, 0)).toContain(JOINED + ' ' + JOINED);
    expect(drawn.checked.fonts.map((font) => font.baseFont).sort()).toEqual(['TeXGyreBonum-Bold', 'TeXGyreBonum-BoldItalic', 'TeXGyreBonum-Italic', 'TeXGyreBonum-Regular']);
    expect(drawn.checked.fonts.every((font) => font.kind === 'cff')).toBe(true);
    // The same document as pictures, and as whatever each page can be.
    const image = await exportWithPages(page, { text: 'image' });
    expect(image.pages.map((entry) => entry.mode)).toEqual(['image', 'image']);
    const pictured = read(image.data);
    expect(pictured.checked.pages.every((sheet) => sheet.images.length === 1 && sheet.texts.every((run) => run.mode === 3))).toBe(true);
    expect(words(pictured.checked, 0)).toContain(TURKISH);
    const auto = await exportWithPages(page, {});
    expect(auto.pages.map((entry) => entry.mode), 'the default writes what can be drawn as drawings').toEqual(['vector', 'vector']);
    if (drawn.poppler) {
      expect(drawn.poppler.fonts.map((font) => [font.type, font.embedded, font.unicode])).toEqual([0, 1, 2, 3].map(() => ['CID Type 0C', true, true]));
      expect(drawn.poppler.images).toEqual([]);
      const lines = drawn.poppler.text[0].split('\n').map((entry) => entry.trim()).filter((entry) => entry !== '');
      expect(lines.slice(0, 3)).toEqual(['Bölüm 1', TURKISH, JOINED + ' ' + JOINED]);
      // The text stands where the picture has it: the invisible text of the picture was placed by the same layout.
      for (const number of [1, 2]) {
        const drawnWords = popplerWords(drawn.bytes).words.filter((word) => word.page === number);
        const picturedWords = popplerWords(pictured.bytes).words.filter((word) => word.page === number);
        expectSameWords(drawnWords, picturedWords, `page ${number}`);
        expect(Math.abs(Math.min(...drawnWords.map((word) => word.xMin)) - 36), `page ${number}: the left margin is half an inch`).toBeLessThanOrEqual(1.5);
        // And the two files look alike when Poppler paints them: the same ink in every band of rows and of
        // columns, and the borders and the shading of the paragraph where the picture has them.
        const drawnRaster = popplerPage(drawn.bytes, number, 100);
        const picturedRaster = popplerPage(pictured.bytes, number, 100);
        expect(inkDifferences(drawnRaster, picturedRaster), `page ${number}: where the drawn page and its picture differ`).toEqual([]);
        expectSameMarks(drawnRaster, picturedRaster, `page ${number}`);
      }
    }
    await expectResponsive(page);
  });

  test('ten pages of text are much smaller as text than as pictures', async ({ page }, info) => {
    test.setTimeout(240_000);
    await openDocPage(page);
    await showDoc(page, await fontDocx(10, ABSENT));
    await withPackageFont(page);
    const vector = await made(page, { text: 'vector' });
    const image = await made(page, { text: 'image' });
    expect(vector.checked.pages.length).toBe(10);
    const sizes = { vector: vector.bytes.length, image: image.bytes.length };
    await info.attach('pdf-vector-sizes', { body: JSON.stringify(sizes), contentType: 'application/json' });
    // The four faces of the font go in whole (a font with CFF outlines is not cut down: about 210 KB together,
    // compressed), whatever the number of pages; the pages themselves are about 2 KB each, where the picture of
    // such a page is about 60 KB. Measured in Chromium: 228 KB drawn, 626 KB as pictures. With ten pages of a few
    // lines the fonts are most of the file, so the drawn file is less than half of the other one, not a tenth.
    expect(sizes.vector, `ten drawn pages have ${sizes.vector} bytes`).toBeLessThan(300 * 1024);
    expect(sizes.vector * 2, `drawn ${sizes.vector} bytes, pictured ${sizes.image} bytes`).toBeLessThan(sizes.image);
  });

  test('a page in a font of the device stays a picture, and "vector" names it', async ({ page }) => {
    await openDocPage(page);
    // The second page has a line in Arial.
    await showDoc(page, await fontDocx(3, ABSENT, (number) => (number === 2 ? para(inFont('Bu satır Arial ile', 'Arial')) : '')));
    await withPackageFont(page);
    const auto = await exportWithPages(page, {});
    expect(auto.error).toBeUndefined();
    expect(auto.pages.map((entry) => entry.mode)).toEqual(['vector', 'image', 'vector']);
    expect(auto.pages[1].reason).toMatch(/font of the device/);
    const mixed = read(auto.data);
    expect(mixed.checked.pages.map((sheet) => sheet.images.length)).toEqual([0, 1, 0]);
    expect(mixed.checked.pages[1].texts.every((run) => run.mode === 3), 'the pictured page has its text lying over it').toBe(true);
    expect(words(mixed.checked, 1)).toContain('Bu satır Arial ile');
    expect(mixed.checked.fonts.map((font) => font.baseFont)).toContain('BYMGlyphless');
    const refused = await exportWithPages(page, { text: 'vector' });
    expect(refused.data).toBeUndefined();
    expect(refused.error).toMatchObject({ name: 'PdfExportError', code: 'unsupported', page: 2 });
    expect((refused.error as any).reason).toMatch(/font of the device/);
    expect(refused.error!.message).toMatch(/^Page 2 cannot be written as text and drawings/);
    // A character the font does not have (the lira sign) is painted with another font by the browser.
    await showDoc(page, await fontDocx(1, ABSENT, () => para(inFont('Tutar: 100 ₺', ABSENT))));
    await withPackageFont(page);
    const lira = await exportWithPages(page, {});
    expect(lira.pages.map((entry) => entry.mode)).toEqual(['image']);
    expect(lira.pages[0].reason).toMatch(/no glyph for U\+20BA/);
    await expectResponsive(page);
  });

  test('a registered TrueType font is cut down to the glyphs that are used; a picture of the page is a picture of its own', async ({ page }) => {
    await page.goto('/e2e/pages/doc.html');
    await expect(page.locator('#state')).toHaveText('ready');
    const font = testFont();
    await page.evaluate(([bytes, family]) => (window as any).BymOffice.fonts.registerFont({ family, source: new Uint8Array(bytes as number[]) }), [Array.from(font), BOXES] as [number[], string]);
    const png = photoPng(240, 160);
    // 240 by 160 px: 180 by 120 points.
    const body = para(inFont('ABC abc', BOXES)) + para(picture(240, 160)) + para(inFont('xyz', BOXES));
    await showDoc(page, await buildDocx({ styles: fontStyles(BOXES, 24), body: body + docSection(), rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', png]] }));
    await expect.poll(() => page.evaluate(() => (window as any).__view.layoutDone), { timeout: 60_000 }).toBe(true);
    const out = await exportWithPages(page, { text: 'vector' });
    expect(out.error).toBeUndefined();
    const pdf = read(out.data);
    expect(pdf.checked.fonts.length).toBe(1);
    expect(pdf.checked.fonts[0].kind).toBe('truetype');
    expect(pdf.checked.fonts[0].baseFont).toMatch(/^[A-Z]{6}\+/);
    expect(pdf.checked.fonts[0].file.length, 'the subset is smaller than the font').toBeLessThan(font.length);
    expect(words(pdf.checked, 0)).toBe('ABC abcxyz');
    // Every glyph of the font is one em wide: the first line is seven ems long.
    const first = pdf.checked.pages[0].texts[0];
    expect(near(first.advance, 7 * 12, 0.05), `seven characters of 12 pt are ${first.advance} points wide`).toBe(true);
    // The photograph is a picture of its own, at its place, not the whole page.
    expect(pdf.checked.pages[0].images.length).toBe(1);
    const photo = pdf.checked.pages[0].images[0];
    expect(photo.width * photo.height, 'far fewer pixels than a page has').toBeLessThan(1240 * 1754 / 4);
    expect(near(photo.matrix![0], 180, 0.5) && near(photo.matrix![3], 120, 0.5), `the picture is 180 by 120 points (${photo.matrix})`).toBe(true);
    if (pdf.poppler) {
      expect(pdf.poppler.fonts.map((entry) => [entry.type, entry.embedded, entry.subset, entry.unicode])).toEqual([['CID TrueType', true, true, true]]);
      expect(pdf.poppler.images.length).toBe(1);
    }
    await page.evaluate((family) => (window as any).BymOffice.fonts.unregisterFont(family), BOXES);
  });

  test('a sheet whose font was registered is written as its cells, lines and text; the bridge says how many pages were drawn', async ({ page }) => {
    await page.goto('/e2e/pages/view.html');
    await expect(page.locator('#state')).toHaveText('ready');
    const font = testFont();
    // The font stands in for the font of the cells wherever the device has it or not.
    await page.evaluate(([bytes, family]) => (window as any).BymOffice.fonts.registerFont({ family, source: new Uint8Array(bytes as number[]), as: ['Calibri'], prefer: true }), [Array.from(font), BOXES] as [number[], string]);
    const data = `<sheetData><row r="1">${str('A1', 'Urun')}${str('B1', 'Tutar')}</row><row r="2">${str('A2', 'Cay')}${num('B2', 125)}</row><row r="3">${str('A3', 'Kahve')}${num('B3', 40)}</row><row r="4">${str('A4', 'ABCDEFGHIJKL')}${num('B4', 7)}</row></sheetData><printOptions gridLines="1"/>${A4_SETUP}`;
    await page.evaluate(([input, opts]) => (window as any).__mountView(input, opts), [Array.from(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(data) }] })), { locale: 'en-US' }] as [number[], object]);
    await expect(page.locator('.bym-canvas')).toBeVisible();
    const out = await exportWithPages(page, { text: 'vector' });
    expect(out.error, 'the page of the sheet can be drawn').toBeUndefined();
    expect(out.pages).toEqual([{ page: 1, mode: 'vector', reason: undefined }]);
    const pdf = read(out.data);
    expect(pdf.checked.pages[0].images.length).toBe(0);
    const cells = pdf.checked.pages[0].texts.map((run) => run.text).sort();
    // The long text of A4 is cut off by the cell beside it: only the letters that show are text of the page.
    const cut = cells.filter((text) => text.charAt(0) === 'A')[0];
    expect(cut.length >= 2 && cut.length < 12 && 'ABCDEFGHIJKL'.indexOf(cut) === 0, `what is written of the text that is cut off: "${cut}"`).toBe(true);
    expect(pdf.checked.pages[0].text).not.toContain('L');
    expect(cells.filter((text) => text !== cut)).toEqual(['125', '40', '7', 'Cay', 'Kahve', 'Tutar', 'Urun']);
    expect(pdf.checked.pages[0].shapes.some((shape) => shape.operator === 'S' || shape.operator === 'f'), 'the grid lines are drawings').toBe(true);
    const image = await exportWithPages(page, { text: 'image' });
    const pictured = read(image.data);
    // The invisible text of the pictured page holds the same letters of it, and no others.
    expect(pictured.checked.pages[0].text).toContain(cut);
    expect(pictured.checked.pages[0].text).not.toContain('L');
    if (pdf.poppler) {
      expect(inkDifferences(popplerPage(pdf.bytes, 1, 100), popplerPage(pictured.bytes, 1, 100)), 'where the drawn sheet and its picture differ').toEqual([]);
      // "Tutar" does not fit its cell: its last letter shows in part, and is 11 points wide.
      expectSameWords(popplerWords(pdf.bytes).words, popplerWords(pictured.bytes).words, 'the sheet', 11);
    }
    await page.evaluate((family) => (window as any).BymOffice.fonts.unregisterFont(family), BOXES);
  });
});
