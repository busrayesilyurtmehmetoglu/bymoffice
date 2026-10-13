// Making a PDF of a view, the parts that need no browser: the options, the
// limits, the picture of a page as a canvas gives it, the name of the file.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocumentPrintError, planPrint } from '../src/ui/doc-print';
import {
  DEFAULT_PDF_DPI,
  DEFAULT_PDF_QUALITY,
  MAX_PDF_DPI,
  MIN_PDF_DPI,
  PdfExportError,
  canvasPageImage,
  checkCancelled,
  pdfFileName,
  pdfLimits,
  pdfSettings,
  toPdfExportError,
} from '../src/pdf/export';
import { PdfError } from '../src/pdf/writer';

test('the options of a PDF have defaults and limits', () => {
  assert.deepEqual(pdfSettings({}), { dpi: 150, image: 'auto', quality: 0.85 });
  assert.equal(pdfSettings({ image: 'jpeg' }).image, 'jpeg');
  assert.deepEqual([DEFAULT_PDF_DPI, MIN_PDF_DPI, MAX_PDF_DPI, DEFAULT_PDF_QUALITY], [150, 72, 300, 0.85]);
  assert.deepEqual(pdfSettings({ dpi: 300, image: 'lossless', quality: 0.5 }), { dpi: 300, image: 'lossless', quality: 0.5 });
  assert.equal(pdfSettings({ dpi: 10 }).dpi, 72, 'a resolution below the range is the lowest');
  assert.equal(pdfSettings({ dpi: 1200 }).dpi, 300, 'one above it the highest');
  assert.equal(pdfSettings({ quality: 1 }).quality, 1);
});

test('options that cannot be are refused before any work is done', () => {
  for (const dpi of [0, -1, NaN, Infinity, '150']) assert.throws(() => pdfSettings({ dpi: dpi as number }), RangeError, `dpi ${String(dpi)}`);
  for (const quality of [0, -0.5, 1.01, NaN, '0.5']) assert.throws(() => pdfSettings({ quality: quality as number }), RangeError, `quality ${String(quality)}`);
  assert.throws(() => pdfSettings({ image: 'png' as 'jpeg' }), RangeError);
  assert.throws(() => pdfSettings({ onProgress: 5 as unknown as () => void }), TypeError);
  assert.throws(() => pdfSettings({ signal: 5 as unknown as { aborted: boolean } }), TypeError);
  assert.throws(() => pdfSettings(null as unknown as object), TypeError);
});

test('a long document keeps its resolution: its pages are not held together', () => {
  const a4 = { width: 794, height: 1123 };
  const pages = Array.from({ length: 600 }, () => a4);
  // Printing shares its pixels out over all pages; a PDF holds one page at a time.
  assert.ok(planPrint(pages, { dpi: 150 }).dpi < 150, 'printing lowers the resolution of 600 pages');
  assert.equal(planPrint(pages, { dpi: 150 }, pdfLimits(false)).dpi, 150);
  assert.equal(planPrint(pages, { dpi: 150 }, pdfLimits(true)).dpi, 150);
  // One page is still limited: a phone paints a large page with fewer pixels.
  const large = planPrint([{ width: 5000, height: 5000 }], { dpi: 300 }, pdfLimits(true)).pages[0];
  assert.ok(large.pixelWidth * large.pixelHeight <= pdfLimits(true).maxPagePixels, 'the page stays within the pixels of a canvas');
  assert.ok(pdfLimits(true).maxBytes < pdfLimits(false).maxBytes, 'a phone holds a smaller file');
});

test('errors of the parts become errors of the export with the same reason', () => {
  const error = (source: unknown) => toPdfExportError(source) as PdfExportError;
  assert.ok(error(new DocumentPrintError('empty', 'x')) instanceof PdfExportError, 'a PdfExportError is made');
  assert.equal(error(new DocumentPrintError('empty', 'x')).code, 'empty');
  assert.equal(error(new DocumentPrintError('too-long', 'x')).code, 'too-long');
  assert.equal(error(new DocumentPrintError('cancelled', 'x')).code, 'cancelled');
  assert.equal(error(new DocumentPrintError('unavailable', 'x')).code, 'unavailable');
  assert.equal(error(new PdfError('too-large', 'x')).code, 'too-long');
  assert.equal(error(new PdfError('invalid', 'why')).code, 'failed');
  assert.match(error(new PdfError('invalid', 'why')).message, /why/);
  assert.equal(error(new Error('boom')).code, 'failed');
  assert.equal(error('text').code, 'failed');
  const range = new RangeError('bad option');
  assert.equal(toPdfExportError(range), range, 'a wrong option stays what it is');
  const own = new PdfExportError('cancelled', 'x');
  assert.equal(toPdfExportError(own), own);
  assert.equal(own.name, 'PdfExportError');
});

test('a signal that was aborted cancels', () => {
  assert.doesNotThrow(() => checkCancelled(undefined));
  assert.doesNotThrow(() => checkCancelled({ aborted: false }));
  assert.throws(() => checkCancelled({ aborted: true }), (error: unknown) => error instanceof PdfExportError && error.code === 'cancelled');
});

/** A canvas of the given pixels (RGBA rows) that makes a "JPEG" of them: a frame header with its size and nothing else. */
function fakeCanvas(width: number, height: number, pixel: (x: number, y: number) => [number, number, number], more: { type?: string; toBlob?: boolean; blob?: boolean } = {}) {
  const reads: { y: number; rows: number }[] = [];
  const jpeg = Uint8Array.from([0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, height >> 8, height & 255, width >> 8, width & 255, 3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0, 0xff, 0xd9]);
  const asked: { type: string; quality: number }[] = [];
  const canvas: Record<string, unknown> = {
    width,
    height,
    toDataURL(type: string, quality: number) {
      asked.push({ type, quality });
      return `data:${more.type || 'image/jpeg'};base64,` + Buffer.from(jpeg).toString('base64');
    },
  };
  if (more.toBlob !== false) {
    canvas.toBlob = (callback: (blob: unknown) => void, type: string, quality: number) => {
      asked.push({ type, quality });
      setTimeout(() => callback(more.blob === false ? null : { type: more.type || 'image/jpeg', size: jpeg.length, arrayBuffer: () => Promise.resolve(jpeg.buffer.slice(0)) }), 0);
    };
  }
  const context = {
    getImageData(_x: number, y: number, w: number, rows: number) {
      reads.push({ y, rows });
      const data = new Uint8ClampedArray(w * rows * 4);
      for (let r = 0; r < rows; r++) {
        for (let x = 0; x < w; x++) {
          const colour = pixel(x, y + r);
          data.set([colour[0], colour[1], colour[2], 255], (r * w + x) * 4);
        }
      }
      return { data, width: w, height: rows };
    },
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, context: context as unknown as CanvasRenderingContext2D, reads, asked };
}

const WHITE: [number, number, number] = [255, 255, 255];

test('a page becomes a JPEG of the quality that was asked for', async () => {
  const { canvas, context, asked, reads } = fakeCanvas(40, 2000, (_x, y) => (y === 700 ? [0, 0, 0] : WHITE));
  const picture = await canvasPageImage(canvas, context, { dpi: 150, image: 'jpeg', quality: 0.6 });
  assert.ok(picture.image && picture.image.kind === 'jpeg', 'the picture is a JPEG');
  assert.deepEqual([picture.image.width, picture.image.height, picture.image.components], [40, 2000, 3]);
  assert.deepEqual(asked, [{ type: 'image/jpeg', quality: 0.6 }]);
  // The pixels were only read until the page was known not to be empty: in strips, not all at once.
  assert.ok(reads.length >= 2 && reads.every((read) => read.rows <= 256), 'the pixels are read in strips');
  assert.ok(reads[reads.length - 1].y <= 700 && reads[reads.length - 1].y + reads[reads.length - 1].rows > 700, 'reading stops at the first strip that is not of one colour');
});

test('a page of one colour needs no picture, in either kind', async () => {
  for (const image of ['jpeg', 'lossless'] as const) {
    const { canvas, context, asked } = fakeCanvas(30, 600, () => WHITE);
    assert.deepEqual(await canvasPageImage(canvas, context, { dpi: 150, image, quality: 0.85 }), { fill: [255, 255, 255] });
    assert.deepEqual(asked, [], 'no picture was made');
  }
});

test('a lossless page is its pixels: grey where the page is grey', async () => {
  const grey = fakeCanvas(3, 2, (x, y) => [x * 10 + y, x * 10 + y, x * 10 + y]);
  assert.deepEqual(await canvasPageImage(grey.canvas, grey.context, { dpi: 150, image: 'lossless', quality: 0.85 }), {
    image: { kind: 'samples', data: Uint8Array.from([0, 10, 20, 1, 11, 21]), width: 3, height: 2, colors: 1 },
  });
  const colour = fakeCanvas(2, 1, (x) => [x, 0, 255]);
  const picture = await canvasPageImage(colour.canvas, colour.context, { dpi: 150, image: 'lossless', quality: 0.85 });
  assert.deepEqual(picture, { image: { kind: 'samples', data: Uint8Array.from([0, 0, 255, 1, 0, 255]), width: 2, height: 1, colors: 3 } });
});

test('a canvas without toBlob makes its JPEG through a data address', async () => {
  const { canvas, context } = fakeCanvas(8, 8, (x) => [x, 0, 0], { toBlob: false });
  const picture = await canvasPageImage(canvas, context, { dpi: 150, image: 'jpeg', quality: 0.85 });
  assert.ok(picture.image && picture.image.kind === 'jpeg', 'the picture is a JPEG');
  assert.equal(picture.image.data.length, 23);
});

test('a browser that makes no JPEG gives the pixels instead', async () => {
  // Asked for a JPEG, a canvas may answer with a PNG.
  const { canvas, context } = fakeCanvas(2, 1, (x) => [x, 0, 255], { type: 'image/png' });
  const picture = await canvasPageImage(canvas, context, { dpi: 150, image: 'jpeg', quality: 0.85 });
  assert.deepEqual(picture, { image: { kind: 'samples', data: Uint8Array.from([0, 0, 255, 1, 0, 255]), width: 2, height: 1, colors: 3 } });
});

test('a canvas that gives nothing fails with a reason', async () => {
  const none = fakeCanvas(2, 1, (x) => [x, 0, 255], { blob: false });
  await assert.rejects(canvasPageImage(none.canvas, none.context, { dpi: 150, image: 'jpeg', quality: 0.85 }), (error: unknown) => error instanceof PdfExportError && error.code === 'failed');
  const tainted = fakeCanvas(2, 1, () => WHITE);
  (tainted.context as unknown as { getImageData(): never }).getImageData = () => {
    throw new Error('The canvas has been tainted');
  };
  await assert.rejects(canvasPageImage(tainted.canvas, tainted.context, { dpi: 150, image: 'lossless', quality: 0.85 }), (error: unknown) => error instanceof PdfExportError && error.code === 'failed');
  const empty = fakeCanvas(0, 0, () => WHITE);
  await assert.rejects(canvasPageImage(empty.canvas, empty.context, { dpi: 150, image: 'jpeg', quality: 0.85 }), (error: unknown) => error instanceof PdfExportError && error.code === 'failed');
});

test('the file is named after the document', () => {
  assert.equal(pdfFileName('report.docx', 'document'), 'report.pdf');
  assert.equal(pdfFileName('Ölçüm raporu 2026.xlsx', 'workbook'), 'Ölçüm raporu 2026.pdf');
  assert.equal(pdfFileName('archive.tar.gz', 'document'), 'archive.tar.pdf');
  assert.equal(pdfFileName('no extension', 'document'), 'no extension.pdf');
  assert.equal(pdfFileName('.hidden', 'document'), '.hidden.pdf');
  assert.equal(pdfFileName(undefined, 'document'), 'document.pdf');
  assert.equal(pdfFileName('', 'workbook'), 'workbook.pdf');
  assert.equal(pdfFileName('  ', 'workbook'), 'workbook.pdf');
  // Nothing of a path, and nothing a file system refuses.
  assert.equal(pdfFileName('../../etc/passwd.doc', 'document'), 'passwd.pdf');
  assert.equal(pdfFileName('C:\\Users\\x\\a<b>:"c"|?*.rtf', 'document'), 'a_b___c____.pdf');
  assert.equal(pdfFileName('a\0b\nc.docx', 'document'), 'a_b_c.pdf');
  assert.equal(pdfFileName('x'.repeat(500) + '.docx', 'document').length, 204);
  assert.equal(pdfFileName(12 as unknown as string, 'document'), 'document.pdf');
});
