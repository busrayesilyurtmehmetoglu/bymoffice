// What a review of the PDF writer found untested, each with the test that
// was missing: limits of the document and of its pictures, the predictors of
// lossless pictures, CMYK pictures, the name of the file.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { deflateSync } from 'node:zlib';
import { MIN_PAGE_POINTS, PdfDocument, buildPdf, type PdfPage } from '../src/pdf/document';
import { canvasPageImage, pdfFileName, pdfSettings } from '../src/pdf/export';
import { PdfTextFont } from '../src/pdf/font';
import { changeShare, imageOfPixels, jpegInfo, packSamples, pngFilter, type PdfImage } from '../src/pdf/image';
import { MAX_PAGE_RUNS, MAX_RUN_CHARS, orderRuns, textLayer, type PdfTextRun } from '../src/pdf/text';
import { PdfError, latin1 } from '../src/pdf/writer';
import { checkPdf, parseToUnicode, poppler, popplerPixels, popplerTest, unfilterPng } from './pdf-fixtures';

const text = (bytes: Uint8Array): string => Buffer.from(bytes).toString('latin1');

async function rejection(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    assert.ok(error instanceof PdfError, `a PdfError is thrown, not ${String(error)}`);
    return error.code;
  }
  return 'no error';
}

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof PdfError, `a PdfError is thrown, not ${String(error)}`);
    return error.code;
  }
  return 'no error';
}

const grey = (values: number[], width: number, height: number): PdfImage => ({ kind: 'samples', data: Uint8Array.from(values), width, height, colors: 1 });

// ----------------------------------------------------------------- document

test('a page too small to be one is refused, and the smallest that can be is written', async () => {
  assert.equal(MIN_PAGE_POINTS, 3);
  assert.equal(await rejection(buildPdf([{ width: 0.004, height: 0.004 }])), 'invalid', 'a size that is written as 0');
  assert.equal(await rejection(buildPdf([{ width: 2.99, height: 100 }])), 'invalid');
  assert.equal(await rejection(buildPdf([{ width: 100, height: 1 }])), 'invalid');
  const checked = checkPdf(await buildPdf([{ width: 3, height: 3 }], { date: new Date(0) }));
  assert.deepEqual([checked.pages[0].width, checked.pages[0].height], [3, 3]);
  assert.doesNotMatch(text(await buildPdf([{ width: 3, height: 3 }], { date: new Date(0) })), /MediaBox \[0 0 0 0\]/);
});

test('a page that is added while the one before it is written is refused, and the first one is written', async () => {
  const document = new PdfDocument({ date: new Date(0) });
  const first = document.addPage({ width: 100, height: 100, image: grey([1, 2, 3, 4], 2, 2) });
  assert.equal(await rejection(document.addPage({ width: 100, height: 100 })), 'invalid');
  await first;
  await document.finish();
  assert.equal(document.pageCount, 1);
});

test('after a page that could not be written the file is given up', async () => {
  const document = new PdfDocument({ date: new Date(0) });
  await document.addPage({ width: 100, height: 100 });
  assert.equal(await rejection(document.addPage({ width: 100, height: 100, image: grey([1, 2, 3], 2, 2) })), 'invalid');
  // Half a page may be in the file: nothing more is added to it, and it is not finished.
  assert.equal(await rejection(document.addPage({ width: 100, height: 100 })), 'invalid');
  assert.equal(await rejection(document.finish()), 'invalid');
});

test('a title beyond a thousand characters is cut', async () => {
  const checked = checkPdf(await buildPdf([{ width: 100, height: 100 }], { title: 'T'.repeat(5000), author: 'ş'.repeat(5000), date: new Date(0) }));
  assert.equal(checked.info.Title.length, 1024);
  assert.equal(checked.info.Author, 'ş'.repeat(1024));
});

test('two files that differ in nothing but the content of a stream have different identifiers', async () => {
  // Two pictures whose compressed bytes are equally long: the size of the file is the same.
  const one = await buildPdf([{ width: 100, height: 100, image: grey([10, 20, 30, 40], 2, 2) }], { date: new Date(0) });
  const two = await buildPdf([{ width: 100, height: 100, image: grey([40, 30, 20, 10], 2, 2) }], { date: new Date(0) });
  assert.equal(one.length, two.length, 'the two files are equally long');
  const idOf = (file: Uint8Array) => (/\/ID \[<([0-9a-f]{32})> <\1>\]/.exec(text(file)) as RegExpExecArray)[1];
  assert.notEqual(idOf(one), idOf(two));
  // The same for the text of a page: one run, at one place or at another.
  const at = (x: number) => buildPdf([{ width: 100, height: 100, text: { width: 100, runs: [{ text: 'a ', x, y: 50, width: 10, size: 10 }] } }], { date: new Date(0) });
  const left = await at(13);
  const right = await at(31);
  assert.equal(left.length, right.length, 'the two files are equally long');
  assert.notEqual(idOf(left), idOf(right));
});

// ---------------------------------------------------------------- pictures

test('pixels with as much red as green and another blue are colour, not grey', () => {
  const rgba = Uint8Array.from([10, 10, 10, 255, 20, 20, 200, 255]);
  assert.deepEqual(imageOfPixels(rgba, 2, 1), { image: { kind: 'samples', data: Uint8Array.from([10, 10, 10, 20, 20, 200]), width: 2, height: 1, colors: 3 } });
  // And as much green as blue with another red.
  assert.equal((imageOfPixels(Uint8Array.from([10, 10, 10, 255, 200, 20, 20, 255]), 2, 1).image as { colors: number }).colors, 3);
});

/** A JPEG of the given segments between its first and its last marker. */
const jpegOf = (...segments: number[][]): Uint8Array => Uint8Array.from([0xff, 0xd8, ...segments.reduce((all, segment) => all.concat(segment), [] as number[]), 0xff, 0xd9]);
const frame = (marker: number, width: number, height: number, components: number, bits = 8): number[] => {
  const out = [0xff, marker, 0, 8 + components * 3, bits, height >> 8, height & 255, width >> 8, width & 255, components];
  for (let i = 0; i < components; i++) out.push(i + 1, 0x11, 0);
  return out;
};

test('the frame header of a JPEG is not taken from a table that looks like one', () => {
  // A Huffman table (0xC4) and an arithmetic one (0xCC) whose bytes read as a size of 0x0808 x 0x0808 in three components.
  const table = (marker: number): number[] => [0xff, marker, 0, 17, 8, 8, 8, 8, 8, 3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0];
  assert.deepEqual(jpegInfo(jpegOf(table(0xc4), frame(0xc0, 40, 30, 3))), { width: 40, height: 30, components: 3, adobe: false });
  assert.deepEqual(jpegInfo(jpegOf(table(0xcc), frame(0xc2, 40, 30, 1))), { width: 40, height: 30, components: 1, adobe: false });
  assert.equal(codeOf(() => jpegInfo(jpegOf(table(0xc8)))), 'invalid', '0xC8 is no frame header either');
});

test('a JPEG with samples of more than eight bits, or cut inside its frame header, is refused', () => {
  assert.equal(codeOf(() => jpegInfo(jpegOf(frame(0xc0, 40, 30, 3, 12)))), 'invalid');
  // The header says 17 bytes and the file ends after the number of components.
  const whole = jpegOf(frame(0xc0, 40, 30, 3));
  assert.equal(codeOf(() => jpegInfo(whole.subarray(0, 12))), 'invalid');
  assert.equal(codeOf(() => jpegInfo(whole.subarray(0, 14))), 'invalid');
});

/** The marker Adobe's programs write into a JPEG; `transform` 0 says that four components are CMYK as they are. */
const ADOBE = [0xff, 0xee, 0, 14, 0x41, 0x64, 0x6f, 0x62, 0x65, 0, 100, 0, 0, 0, 0, 0];

/**
 * A JPEG of 8 x 8 pixels in four components, written by hand: one block per
 * component with nothing but its mean, which is 255 for the first component
 * and 0 for the other three.
 */
function cmykJpeg(adobe: boolean): Uint8Array {
  const bytes: number[] = [];
  if (adobe) bytes.push(...ADOBE);
  // One quantisation table: the mean in steps of 8, everything else in steps of 1.
  bytes.push(0xff, 0xdb, 0, 67, 0, 8, ...new Array(63).fill(1));
  bytes.push(...frame(0xc0, 8, 8, 4));
  // The table of the means: "0" is followed by 7 bits (category 7), "10" by 8 (category 8).
  bytes.push(0xff, 0xc4, 0, 21, 0x00, 1, 1, ...new Array(14).fill(0), 7, 8);
  // The table of the rest: "0" ends the block.
  bytes.push(0xff, 0xc4, 0, 20, 0x10, 1, ...new Array(15).fill(0), 0);
  bytes.push(0xff, 0xda, 0, 14, 4, 1, 0x00, 2, 0x00, 3, 0x00, 4, 0x00, 0, 63, 0);
  // First component: "0", +127 ("1111111"), end "0". The others: "10", -128 ("01111111"), end "0". Padded with ones.
  const bits = '0' + '1111111' + '0' + ('10' + '01111111' + '0').repeat(3) + '111111';
  for (let i = 0; i < bits.length; i += 8) {
    const value = parseInt(bits.substr(i, 8), 2);
    bytes.push(value);
    // A byte of ones is followed by a zero, so that it is not taken for a marker.
    if (value === 0xff) bytes.push(0);
  }
  return jpegOf(bytes);
}

test('a JPEG of four components is CMYK, and one of Adobe is stored the other way round', async () => {
  const plain = cmykJpeg(false);
  const adobe = cmykJpeg(true);
  assert.deepEqual(jpegInfo(plain), { width: 8, height: 8, components: 4, adobe: false });
  assert.deepEqual(jpegInfo(adobe), { width: 8, height: 8, components: 4, adobe: true });
  const file = await buildPdf(
    [
      { width: 80, height: 80, image: { kind: 'jpeg', data: plain, width: 8, height: 8, components: 4 } },
      { width: 80, height: 80, image: { kind: 'jpeg', data: adobe, width: 8, height: 8, components: 4 } },
    ],
    { date: new Date(0) },
  );
  const checked = checkPdf(file);
  assert.equal(checked.pages[0].images[0].colorSpace, 'DeviceCMYK');
  assert.equal(checked.pages[0].images[0].decode, undefined);
  assert.equal(checked.pages[1].images[0].colorSpace, 'DeviceCMYK');
  // Adobe writes 255 for "no ink": the values are turned around when they are read.
  assert.deepEqual(checked.pages[1].images[0].decode, [1, 0, 1, 0, 1, 0, 1, 0]);
});

popplerTest('Poppler paints a CMYK JPEG of Adobe with its inks turned around', async () => {
  const page = (adobe: boolean): PdfPage => ({ width: 80, height: 80, image: { kind: 'jpeg', data: cmykJpeg(adobe), width: 8, height: 8, components: 4 } });
  // Stored: all of the first component, none of the others. As it is: cyan and nothing else.
  const plain = await buildPdf([page(false)], { date: new Date(0) });
  assert.equal(poppler(plain).complaints, '');
  const cyan = popplerPixels(plain).at(40, 40);
  assert.ok(cyan[0] < 100 && cyan[1] > 120 && cyan[2] > 150, `a picture without the marker is cyan (${cyan})`);
  // Turned around: no cyan and all of magenta, yellow and black.
  const adobe = await buildPdf([page(true)], { date: new Date(0) });
  assert.equal(poppler(adobe).complaints, '');
  const dark = popplerPixels(adobe).at(40, 40);
  assert.ok(dark[0] < 80 && dark[1] < 80 && dark[2] < 80, `a picture of Adobe is dark (${dark})`);
});

// --------------------------------------------------------------- predictors

test('every row of a lossless picture is stored by the filter that leaves the least', () => {
  // Row 0 is flat (every pixel like the one before it), row 1 repeats row 0, row 2 is a ramp, row 3 is noise.
  const rows = [
    [200, 200, 200, 200, 200, 200],
    [200, 200, 200, 200, 200, 200],
    [10, 20, 30, 40, 50, 60],
    [7, 250, 3, 190, 66, 121],
  ];
  const samples = Uint8Array.from(rows.reduce((all, row) => all.concat(row), [] as number[]));
  const filtered = pngFilter(samples, 6, 4, 1);
  assert.equal(filtered.length, 7 * 4);
  // Sub, Up, Sub (a ramp is a constant step), and for the noise whatever leaves the least.
  assert.deepEqual([filtered[0], filtered[7], filtered[14]], [1, 2, 1]);
  assert.deepEqual(Array.from(filtered.subarray(8, 14)), [0, 0, 0, 0, 0, 0], 'a row like the one above it is all zeros');
  assert.deepEqual(Array.from(filtered.subarray(15, 21)), [10, 10, 10, 10, 10, 10]);
  assert.deepEqual(Array.from(unfilterPng(filtered, 6, 4, 1)), Array.from(samples), 'undoing the filters gives the samples back');
  // A picture whose left half changes from row to row and whose right half from pixel to pixel: only the filter that
  // looks at both neighbours (Paeth) leaves nothing, and what it leaves is undone to the same samples.
  const both = new Uint8Array(8 * 4);
  for (let y = 0; y < 4; y++) for (let x = 0; x < 8; x++) both[y * 8 + x] = x < 4 ? 50 + 10 * y : 20 * x;
  const twoWays = pngFilter(both, 8, 4, 1);
  assert.deepEqual([twoWays[9], twoWays[18], twoWays[27]], [4, 4, 4]);
  assert.deepEqual(Array.from(unfilterPng(twoWays, 8, 4, 1)), Array.from(both));
  // Colour: the pixel before is three bytes back.
  const colour = new Uint8Array(5 * 3 * 3);
  for (let i = 0; i < colour.length; i++) colour[i] = (i * 53 + (i % 3) * 91) & 255;
  assert.deepEqual(Array.from(unfilterPng(pngFilter(colour, 5, 3, 3), 5, 3, 3)), Array.from(colour));
  assert.equal(codeOf(() => pngFilter(new Uint8Array(5), 2, 2, 3)), 'invalid');
});

test('a lossless picture is written with its predictor, and read back as it was', async () => {
  const width = 40;
  const height = 30;
  const samples = new Uint8Array(width * height * 3);
  // A soft ramp in three channels: what the filters are good at.
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) samples.set([x * 6, y * 8, (x + y) * 3], (y * width + x) * 3);
  const file = await buildPdf([{ width: 400, height: 300, image: { kind: 'samples', data: samples, width, height, colors: 3 } }], { date: new Date(0) });
  const image = checkPdf(file).pages[0].images[0];
  assert.equal(image.predictor, true);
  assert.deepEqual(Array.from(image.samples as Uint8Array), Array.from(samples));
  assert.match(text(file), /\/DecodeParms <<\/Predictor 15 \/Colors 3 \/BitsPerComponent 8 \/Columns 40>>/);
  assert.ok(image.data.length < deflateSync(samples, { level: 9 }).length, `the predictor makes the ramp smaller (${image.data.length} bytes) than it is without one`);
  // A picture that was packed before it was handed over is written as it is.
  const packed = await packSamples({ kind: 'samples', data: samples, width, height, colors: 3 });
  assert.equal(packed.kind, 'packed');
  const again = checkPdf(await buildPdf([{ width: 400, height: 300, image: packed }], { date: new Date(0) })).pages[0].images[0];
  assert.deepEqual(Array.from(again.data), Array.from(packed.data));
  assert.deepEqual(Array.from(again.samples as Uint8Array), Array.from(samples));
  assert.equal(await rejection(buildPdf([{ width: 10, height: 10, image: { kind: 'packed', data: new Uint8Array(0), width: 2, height: 2, colors: 2 as 1, predictor: true } }])), 'invalid');
});

test('a picture of text is stored without predictors, which only make it larger', async () => {
  // White paper with dark strokes: most pixels are like the one before them.
  const width = 120;
  const height = 96;
  const samples = new Uint8Array(width * height);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) samples[y * width + x] = textPixel(x, y)[0];
  const image: PdfImage = { kind: 'samples', data: samples, width, height, colors: 1 };
  const packed = await packSamples(image as PdfImage & { kind: 'samples' });
  assert.equal(packed.predictor, false);
  const file = await buildPdf([{ width: 400, height: 300, image }], { date: new Date(0) });
  assert.doesNotMatch(text(file), /DecodeParms/);
  const read = checkPdf(file).pages[0].images[0];
  assert.equal(read.predictor, undefined);
  assert.deepEqual(Array.from(read.samples as Uint8Array), Array.from(samples));
  // Asked for, either way is taken.
  assert.equal((await packSamples(image as PdfImage & { kind: 'samples' }, true, true)).predictor, true);
});

popplerTest('Poppler reads a picture that is stored with predictors', async () => {
  // A ramp of red on the left half and of blue on the right: soft enough for the predictors to be used.
  const samples = new Uint8Array(16 * 16 * 3);
  for (let y = 0; y < 16; y++) for (let x = 0; x < 16; x++) samples.set(x < 8 ? [255 - x * 10, 0, 0] : [0, 0, 255 - x * 10], (y * 16 + x) * 3);
  const file = await buildPdf([{ width: 160, height: 160, image: { kind: 'samples', data: samples, width: 16, height: 16, colors: 3 } }], { date: new Date(0) });
  assert.equal(checkPdf(file).pages[0].images[0].predictor, true);
  assert.equal(poppler(file).complaints, '');
  const pixels = popplerPixels(file);
  // Columns 4 and 12 of 16.
  assert.deepEqual(pixels.at(45, 80), [215, 0, 0]);
  assert.deepEqual(pixels.at(125, 80), [0, 0, 135]);
});

// --------------------------------------------------------------------- auto

/** The pixels of a page as a canvas has them: RGBA. */
function pixelsOf(width: number, height: number, pixel: (x: number, y: number) => [number, number, number]): Uint8Array {
  const rgba = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) rgba.set([...pixel(x, y), 255], (y * width + x) * 4);
  return rgba;
}

/** A page of "text": white, with short dark strokes on some lines. */
const textPixel = (x: number, y: number): [number, number, number] => (y % 12 < 6 && (x * 7 + y * 3) % 9 === 0 ? [20, 20, 20] : [255, 255, 255]);
/** A "photograph": no pixel like the one before it. */
const photoPixel = (x: number, y: number): [number, number, number] => [(x * 37 + y * 11) & 255, (x * 17 + y * 29 + ((x * y) % 7) * 13) & 255, (x * 5 + y * 43) & 255];

test('how much of a page changes from pixel to pixel tells text from photographs', () => {
  const share = (pixel: (x: number, y: number) => [number, number, number]) => changeShare(imageOfPixels(pixelsOf(120, 96, pixel), 120, 96).image as PdfImage & { kind: 'samples' });
  assert.ok(share(textPixel) < 0.25, `a page of text changes little (${share(textPixel)})`);
  assert.ok(share(photoPixel) > 0.6, `a photograph changes at nearly every pixel (${share(photoPixel)})`);
  // Half a page of each lies in between.
  const half = share((x, y) => (y < 48 ? textPixel(x, y) : photoPixel(x, y)));
  assert.ok(half >= 0.25 && half <= 0.6, `half and half lies in between (${half})`);
});

/** A canvas of the given pixels whose JPEG is `jpegBytes` long (a frame header and filling). */
function canvasOf(width: number, height: number, pixel: (x: number, y: number) => [number, number, number], jpegBytes: number) {
  const rgba = pixelsOf(width, height, pixel);
  const jpeg = Uint8Array.from([0xff, 0xd8, ...frame(0xc0, width, height, 3), ...new Array(Math.max(0, jpegBytes - 23)).fill(0), 0xff, 0xd9]);
  let asked = 0;
  const canvas = {
    width,
    height,
    toBlob(callback: (blob: unknown) => void) {
      asked++;
      setTimeout(() => callback({ type: 'image/jpeg', size: jpeg.length, arrayBuffer: () => Promise.resolve(jpeg.buffer.slice(0)) }), 0);
    },
  };
  const context = {
    getImageData(_x: number, y: number, w: number, rows: number) {
      return { data: rgba.subarray(y * w * 4, (y + rows) * w * 4), width: w, height: rows };
    },
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, context: context as unknown as CanvasRenderingContext2D, jpegsMade: () => asked };
}

test('the kind of picture is chosen page by page unless one is asked for', async () => {
  assert.deepEqual(pdfSettings({}), { dpi: 150, image: 'auto', quality: 0.85, text: 'auto', links: true, outline: true });
  assert.equal(pdfSettings({ image: 'auto' }).image, 'auto');
  const auto = { dpi: 150, image: 'auto' as const, quality: 0.85 };
  // A page of text: its pixels, and no JPEG is made at all.
  const page = canvasOf(120, 96, textPixel, 5000);
  const kept = await canvasPageImage(page.canvas, page.context, auto);
  assert.equal(kept.image && kept.image.kind, 'samples');
  assert.equal(page.jpegsMade(), 0);
  // A photograph: a JPEG.
  const photo = canvasOf(120, 96, photoPixel, 5000);
  const taken = await canvasPageImage(photo.canvas, photo.context, auto);
  assert.equal(taken.image && taken.image.kind, 'jpeg');
  // An empty page: neither.
  const empty = canvasOf(120, 96, () => [255, 255, 255], 5000);
  assert.deepEqual(await canvasPageImage(empty.canvas, empty.context, auto), { fill: [255, 255, 255] });
  assert.equal(empty.jpegsMade(), 0);
});

test('for a page that is neither text nor photograph both pictures are made, and the smaller one is taken', async () => {
  const auto = { dpi: 150, image: 'auto' as const, quality: 0.85 };
  const mixed = (x: number, y: number): [number, number, number] => (y < 48 ? textPixel(x, y) : photoPixel(x, y));
  const packedSize = (await packSamples(imageOfPixels(pixelsOf(120, 96, mixed), 120, 96).image as PdfImage & { kind: 'samples' })).data.length;
  // A JPEG that is smaller than the pixels, and one that is larger.
  const small = canvasOf(120, 96, mixed, packedSize - 100);
  const first = await canvasPageImage(small.canvas, small.context, auto);
  assert.equal(first.image && first.image.kind, 'jpeg');
  const large = canvasOf(120, 96, mixed, packedSize + 100);
  const second = await canvasPageImage(large.canvas, large.context, auto);
  assert.equal(second.image && second.image.kind, 'packed');
  assert.equal(large.jpegsMade(), 1);
  assert.equal((second.image as { data: Uint8Array }).data.length, packedSize);
});

// --------------------------------------------------------------------- text

test('a page has no more runs than its limit', () => {
  const runs: PdfTextRun[] = [];
  for (let i = 0; i < MAX_PAGE_RUNS + 5; i++) runs.push({ text: 'x', x: i % 500, y: Math.floor(i / 500), width: 1, size: 1 });
  assert.equal(orderRuns(runs).length, MAX_PAGE_RUNS);
});

test('a very long run is not cut inside a character', () => {
  const font = new PdfTextFont();
  // The pair of surrogates lies across the limit.
  const long = 'x'.repeat(MAX_RUN_CHARS - 1) + '😀' + 'tail ';
  const layer = textLayer([{ text: long, x: 0, y: 10, width: 500, size: 10 }], { height: 100, scale: 1, font });
  const map = parseToUnicode(font.toUnicode());
  const hex = (/<([0-9a-f]+)> Tj/.exec(layer) as RegExpExecArray)[1];
  let written = '';
  for (let i = 0; i < hex.length; i += 4) written += map.get(parseInt(hex.substr(i, 4), 16));
  assert.equal(written, 'x'.repeat(MAX_RUN_CHARS - 1), 'the character that does not fit is left out whole');
  assert.equal(written.indexOf('�'), -1);
});

// --------------------------------------------------------------------- name

test('the name of the download is one every file system takes', () => {
  // A name that is nothing but dots is no name.
  assert.equal(pdfFileName('.', 'document'), 'document.pdf');
  assert.equal(pdfFileName('..', 'document'), 'document.pdf');
  assert.equal(pdfFileName('...', 'document'), 'document.pdf');
  assert.equal(pdfFileName('dir/..', 'workbook'), 'workbook.pdf');
  // Names Windows keeps for devices, with or without an extension, in any case.
  for (const name of ['CON', 'con.docx', 'PRN.xlsx', 'Aux.doc', 'NUL', 'COM1.docx', 'com9', 'LPT1.rtf', 'lpt9.xls']) {
    const made = pdfFileName(name, 'document');
    assert.equal(made, '_' + name.replace(/\.[^.]*$/, '') + '.pdf', name);
  }
  assert.equal(pdfFileName('CON.tar.gz', 'document'), '_CON.tar.pdf', 'a device name before any extension');
  assert.equal(pdfFileName('CONSOLE.docx', 'document'), 'CONSOLE.pdf');
  assert.equal(pdfFileName('COM10.docx', 'document'), 'COM10.pdf');
  // Marks that turn the direction of text around (and so can hide an extension), and controls of the second kind.
  assert.equal(pdfFileName('invoice‮gpj.docx', 'document'), 'invoicegpj.pdf');
  assert.equal(pdfFileName('a‪‫‬‭⁦⁧⁨⁩b\u0085\u009fc.docx', 'document'), 'abc.pdf');
  // At most 200 bytes of UTF-8 before the extension, and no character cut in two.
  const turkish = pdfFileName('ş'.repeat(300) + '.docx', 'document');
  assert.equal(turkish, 'ş'.repeat(100) + '.pdf');
  const emoji = pdfFileName('a' + '😀'.repeat(100) + '.docx', 'document');
  assert.equal(emoji, 'a' + '😀'.repeat(49) + '.pdf', 'one byte and 49 characters of four');
  assert.ok(Buffer.byteLength(emoji.slice(0, -4), 'utf8') <= 200);
  assert.equal(pdfFileName('x'.repeat(500) + '.docx', 'document').length, 204);
});
