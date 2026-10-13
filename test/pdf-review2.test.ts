// What the review of the drawn pages of a PDF found, each as a test: a font
// whose glyphs are broken leaves its page a picture instead of ending the
// file, a subset never grows beyond its font, two fonts never share a name,
// colours CSS does not know are not read, reasons are plain text, and the
// pixels of a picture read for a page that was not drawn are let go of.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cleanReason, createPdfCanvas, parseCssColor, type PdfCanvasMeasure, type PdfPicture } from '../src/pdf/canvas';
import { buildPdf, PdfDocument } from '../src/pdf/document';
import { PdfExportError, pdfSettings, pdfWithFallback } from '../src/pdf/export';
import { PdfFontSet } from '../src/pdf/font-set';
import { FontError, applyLigatures, fontLigatures, readFont, type PdfFontFace } from '../src/pdf/sfnt';
import { MAX_GLYPH_BYTES, glyphUsable, subsetTrueType } from '../src/pdf/subset';
import { canvasMeasure, canvasPictures, drawPage } from '../src/pdf/vector';
import { PdfError } from '../src/pdf/writer';
import { checkPdf, hasPoppler, poppler, popplerTest, popplerWords } from './pdf-fixtures';
import { TEST_GLYPHS, bonum, gsubTable, seeded, sfntFile, testFont, testFontTables, type TestGlyph } from './pdf-font-fixtures';

const DATE = new Date(Date.UTC(2026, 9, 13, 12, 0, 0));

/** Paints `text` in the font on a PDF canvas, as an export does; the page, or why there is none. */
function pageOf(bytes: Uint8Array, text: string): { reason?: string; build?: () => Promise<Uint8Array> } {
  const set = new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes }], () => false);
  const resolved = set.resolve('20px T');
  if (!('face' in resolved)) return { reason: resolved.reason };
  const face = resolved.face;
  const measure: PdfCanvasMeasure = {
    width: (_font, value) => {
      let units = 0;
      for (const char of value) units += face.advance(face.glyph(char.codePointAt(0) as number));
      return (units * 20) / face.unitsPerEm;
    },
  };
  const made = createPdfCanvas({ width: 400, height: 100, pageWidth: 300, pageHeight: 75, fonts: set, measure });
  made.context.font = '20px T';
  made.context.fillText(text, 10, 50);
  if (made.canvas.unsupported !== undefined) return { reason: made.canvas.unsupported };
  const vector = made.canvas.finish();
  return { build: () => buildPdf([{ width: 300, height: 75, vector }], { date: DATE }) };
}

// ---------------------------------------------------------------------------
// B1: a glyph is known to be sound before it is written on a page

test('review: a font whose glyph index points outside its glyphs, or runs backwards, is refused when it is read', () => {
  const beyond = testFontTables();
  // The end of glyph 2 and start of glyph 3 ("B") lies beyond the glyf table.
  beyond.loca[6] = 0xff;
  beyond.loca[7] = 0xf0;
  assert.throws(() => readFont(sfntFile(beyond)), (error: unknown) => error instanceof FontError && /loca/.test(error.message));
  const backwards = testFontTables();
  backwards.loca[6] = 0;
  backwards.loca[7] = 1;
  assert.throws(() => readFont(sfntFile(backwards)), (error: unknown) => error instanceof FontError && /loca/.test(error.message));
  // The end of the last glyph beyond the glyf table: nothing comes after it that would run backwards.
  const last = testFontTables();
  last.loca[last.loca.length - 2] = 0xff;
  last.loca[last.loca.length - 1] = 0xf0;
  assert.throws(() => readFont(sfntFile(last)), (error: unknown) => error instanceof FontError && /loca/.test(error.message));
  // The page of such a font is a picture, whatever of it is written: nothing is found out when the file ends.
  for (const text of ['A', 'AB']) assert.match(String(pageOf(sfntFile(beyond), text).reason), /cannot be embedded.*loca/, text);
});

test('review: a glyph that is broken makes its page a picture, the other glyphs of the font are written', async () => {
  // "D" (glyph 5) names a part the font does not have.
  const glyphs = TEST_GLYPHS.slice();
  glyphs[5] = { advance: 800, parts: [{ glyph: 2, dx: 0, dy: 0 }] };
  const tables = testFontTables({ glyphs });
  const at = ((tables.loca[10] << 8) | tables.loca[11]) * 2;
  tables.glyf[at + 12] = 0x7f;
  tables.glyf[at + 13] = 0xff;
  const bytes = sfntFile(tables);
  const face = readFont(bytes);
  assert.equal(glyphUsable(face, 5), false, 'the glyph with the missing part');
  assert.equal(glyphUsable(face, 2), true, 'a sound glyph of the same font');
  assert.equal(glyphUsable(face, 400), false, 'a glyph the font does not have');
  assert.match(String(pageOf(bytes, 'AD').reason), /glyph 5 of the font BymTest-Regular is not sound/);
  const sound = pageOf(bytes, 'AB');
  assert.equal(sound.reason, undefined);
  checkPdf(await (sound.build as () => Promise<Uint8Array>)());
});

test('review: the outline of a glyph is read to its end before the glyph is used', () => {
  const broken = (change: (glyph: number[]) => void): boolean => {
    const tables = testFontTables();
    // Glyph 2 ("A"): one contour of four points.
    const from = ((tables.loca[4] << 8) | tables.loca[5]) * 2;
    const to = ((tables.loca[6] << 8) | tables.loca[7]) * 2;
    const glyph = tables.glyf.slice(from, to);
    change(glyph);
    for (let i = 0; i < glyph.length; i++) tables.glyf[from + i] = glyph[i];
    return glyphUsable(readFont(sfntFile(tables)), 2);
  };
  assert.equal(broken(() => undefined), true, 'the glyph as it is');
  assert.equal(broken((glyph) => (glyph[1] = 9)), false, 'more contours than the glyph has room for');
  assert.equal(broken((glyph) => ((glyph[10] = 0x40), (glyph[11] = 0))), false, 'a contour that ends beyond the points the glyph holds');
  assert.equal(broken((glyph) => (glyph[13] = 200)), false, 'instructions longer than the glyph');
  // Every point repeats its flag two hundred times: far more points than there are.
  assert.equal(broken((glyph) => ((glyph[16] = 0x09), (glyph[17] = 200))), false, 'flags that repeat beyond the points');
  // Flags that repeat beyond the points and ask for no coordinates at all (the same x and y as the point before).
  assert.equal(broken((glyph) => ((glyph[16] = 0x39), (glyph[17] = 200))), false, 'flags that repeat beyond the points, without coordinates');
  // Two points more than the glyph holds coordinates for: its flags are there (what follows them is read as flags), their coordinates are not.
  assert.equal(broken((glyph) => (glyph[11] = 5)), false, 'more points than coordinates');
  assert.ok(MAX_GLYPH_BYTES >= 65536, 'a glyph of a real font fits the limit on one glyph');
});

test('review: glyphs made of each other in a circle are sound to write; a glyph made of a broken one is not', async () => {
  const circle = TEST_GLYPHS.slice();
  circle[5] = { advance: 800, parts: [{ glyph: 9, dx: 0, dy: 0 }] };
  circle[9] = { advance: 900, parts: [{ glyph: 5, dx: 0, dy: 0 }] };
  const page = pageOf(testFont({ glyphs: circle }), 'D');
  assert.equal(page.reason, undefined);
  checkPdf(await (page.build as () => Promise<Uint8Array>)());
  // "D" is made of "A", and the outline of "A" is cut short.
  const tables = testFontTables();
  const from = ((tables.loca[4] << 8) | tables.loca[5]) * 2;
  tables.glyf[from + 1] = 9;
  assert.equal(glyphUsable(readFont(sfntFile(tables)), 5), false);
  assert.match(String(pageOf(sfntFile(tables), 'D').reason), /not sound/);
});

test('review: a font whose advances stop short of its side bearings is written with what it has', async () => {
  const tables = testFontTables();
  // Two advances for all glyphs, and nothing after them.
  tables.hhea[34] = 0;
  tables.hhea[35] = 2;
  tables.hmtx = tables.hmtx.slice(0, 8);
  const page = pageOf(sfntFile(tables), 'A');
  assert.equal(page.reason, undefined);
  const checked = checkPdf(await (page.build as () => Promise<Uint8Array>)());
  assert.equal(checked.pages[0].text, 'A');
});

test('review: a font that fails when the file is ended ends the file with a reason of its own, and the export is made again as pictures', async () => {
  // The font changes under the file: what was sound when the page was painted is not when the fonts are written.
  const bytes = testFont();
  const set = new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes }], () => false);
  const made = createPdfCanvas({ width: 200, height: 100, pageWidth: 200, pageHeight: 100, fonts: set, measure: { width: () => 14 } });
  made.context.font = '20px T';
  made.context.fillText('A', 10, 50);
  const pdf = new PdfDocument({ date: DATE });
  await pdf.addPage({ width: 200, height: 100, vector: made.canvas.finish() });
  const loca = (readFont(bytes).tables.get('loca') as { offset: number }).offset;
  bytes[loca + 6] = 0xff;
  bytes[loca + 7] = 0xff;
  await assert.rejects(pdf.finish(), (error: unknown) => error instanceof PdfError && error.code === 'font' && /BymTest-Regular/.test(error.message));
  // What a view does with that.
  const asked: string[] = [];
  const out = await pdfWithFallback(pdfSettings({}), async (settings) => {
    asked.push(settings.text);
    if (settings.text !== 'image') throw new PdfError('font', 'The font X could not be written');
    return 'pictures';
  });
  assert.equal(out, 'pictures');
  assert.deepEqual(asked, ['auto', 'image'], 'made once more, every page a picture');
  await assert.rejects(
    pdfWithFallback(pdfSettings({ text: 'vector' }), async () => {
      throw new PdfError('font', 'The font X could not be written');
    }),
    (error: unknown) => error instanceof PdfExportError && error.code === 'unsupported' && /font X/.test(error.message),
  );
  // Anything else is what it was, and pictures that fail are not made again.
  await assert.rejects(pdfWithFallback(pdfSettings({}), async () => Promise.reject(new PdfError('too-large', 'too large'))), (error: unknown) => error instanceof PdfError && error.code === 'too-large');
  let times = 0;
  await assert.rejects(
    pdfWithFallback(pdfSettings({ text: 'image' }), async () => {
      times++;
      throw new PdfError('font', 'never from pictures');
    }),
    (error: unknown) => error instanceof PdfError,
  );
  assert.equal(times, 1);
});

test('review: with fonts broken at random a page is a sound PDF or a picture, and the file is never lost when it is ended', async () => {
  const random = seeded(4242);
  let written = 0;
  let pictures = 0;
  let asked = 0;
  for (let round = 0; round < 500; round++) {
    const tables = testFontTables({ extra: round % 3 === 0 ? { 'cvt ': [0, 1], prep: [0xb0, 1] } : undefined });
    // The tables a subset is made of, a few bytes each time.
    const tags = ['glyf', 'loca', 'hmtx', 'hhea', 'maxp', 'head'];
    const changes = 1 + Math.floor(random() * 4);
    for (let i = 0; i < changes; i++) {
      const table = tables[tags[Math.floor(random() * (round % 2 === 0 ? 2 : tags.length))]];
      const at = Math.floor(random() * table.length);
      table[at] = random() < 0.3 ? 0xff : random() < 0.5 ? 0 : Math.floor(random() * 256);
    }
    if (random() < 0.1) tables.hmtx = tables.hmtx.slice(0, Math.floor(random() * tables.hmtx.length));
    if (random() < 0.1) tables.glyf = tables.glyf.slice(0, Math.floor(random() * tables.glyf.length));
    const page = pageOf(sfntFile(tables), 'ABCD şİğ');
    if (!page.build) {
      assert.equal(typeof page.reason, 'string', `round ${round}: a page that is not drawn says why`);
      pictures++;
      continue;
    }
    let file: Uint8Array;
    try {
      file = await page.build();
    } catch (error) {
      assert.fail(`round ${round}: the file was lost when it was ended: ${String(error)}`);
    }
    try {
      checkPdf(file);
    } catch (error) {
      assert.fail(`round ${round}: the file is not sound: ${String(error)}`);
    }
    written++;
    // Poppler reads some of them: the structure of the file and of its font have nothing to complain of.
    if (hasPoppler() && asked < 40) {
      asked++;
      const report = poppler(file);
      assert.equal(report.complaints, '', `round ${round}: Poppler has nothing to say against the file`);
      assert.equal(report.fonts.length, 1, `round ${round}: Poppler finds the font`);
    }
  }
  assert.ok(written > 50 && pictures > 50, `both happen (${written} written, ${pictures} pictures)`);
});

// ---------------------------------------------------------------------------
// B2: a subset is never larger than what it was cut from

test('review: a subset is no larger than the glyphs of its font, whatever the font says of where its glyphs are', () => {
  // Glyphs that all claim the same 60,000 bytes: read one after the other they would be in the subset once each.
  const tables = testFontTables();
  const glyf = tables.glyf.concat(new Array(120000 - tables.glyf.length).fill(0));
  const overlapping: number[] = [];
  for (let glyph = 0; glyph <= TEST_GLYPHS.length; glyph++) {
    const value = glyph % 2 === 1 ? 60000 : 0;
    overlapping.push(value >> 8, value & 255);
  }
  assert.throws(() => readFont(sfntFile({ ...tables, glyf, loca: overlapping })), (error: unknown) => error instanceof FontError && /loca/.test(error.message));
  // A font whose glyphs follow each other, one of them very large: the subset holds it once.
  const big: TestGlyph[] = TEST_GLYPHS.slice();
  const contour: [number, number][] = [];
  for (let i = 0; i < 3000; i++) contour.push([i % 2 === 0 ? 0 : 500, i]);
  big[9] = { advance: 900, contours: [contour] };
  const bytes = testFont({ glyphs: big });
  const face = readFont(bytes);
  const subset = subsetTrueType(face, [2, 9]);
  const glyfLength = (face.tables.get('glyf') as { length: number }).length;
  assert.ok(subset.length <= glyfLength + 4096, `the subset (${subset.length}) is no larger than the glyphs of the font (${glyfLength}) and its small tables`);
  // One glyph beyond the limit on a glyph is not written at all.
  const huge = testFontTables();
  const room = MAX_GLYPH_BYTES + 64;
  const padded = huge.glyf.concat(new Array(room).fill(0));
  // The last glyph (9) ends where the padding does: long loca entries are in words, so the format is changed to long.
  const long: number[] = [];
  for (let glyph = 0; glyph <= TEST_GLYPHS.length; glyph++) {
    const value = glyph === TEST_GLYPHS.length ? padded.length : ((huge.loca[glyph * 2] << 8) | huge.loca[glyph * 2 + 1]) * 2;
    long.push((value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255);
  }
  huge.head[51] = 1;
  const hugeFace = readFont(sfntFile({ ...huge, glyf: padded, loca: long }));
  assert.equal(glyphUsable(hugeFace, 9), false, 'a glyph larger than a glyph may be');
  assert.equal(glyphUsable(hugeFace, 2), true);
});

// ---------------------------------------------------------------------------
// B6: fonts of one file have names of their own

test('review: two fonts with one name are two names in the file, with the same glyphs or not', async () => {
  // The same name, other outlines.
  const first = testFont();
  const other = TEST_GLYPHS.slice();
  other[2] = { advance: 650, contours: [[[0, 0], [0, 100], [100, 100], [100, 0]]] };
  const second = testFont({ glyphs: other });
  const whole = (bytes: Uint8Array, family: string) => ({ family, weight: 400, style: 'normal' as const, bytes });
  for (const fsType of [0, 0x0100]) {
    const a = testFont({ fsType });
    const b = testFont({ fsType, glyphs: other });
    const set = new PdfFontSet([whole(a, 'One'), whole(b, 'Two'), whole(testFont({ fsType, glyphs: other, weight: 500 }), 'Three')], () => false);
    const widths: { [family: string]: number } = { One: 14, Two: 13, Three: 13 };
    const made = createPdfCanvas({ width: 200, height: 100, pageWidth: 200, pageHeight: 100, fonts: set, measure: { width: (font) => widths[font.replace(/^20px /, '')] } });
    for (const family of ['One', 'Two', 'Three']) {
      made.context.font = `20px ${family}`;
      made.context.fillText('A', 10, 50);
    }
    assert.equal(made.canvas.unsupported, undefined);
    const checked = checkPdf(await buildPdf([{ width: 200, height: 100, vector: made.canvas.finish() }], { date: DATE }));
    const names = checked.fonts.map((font) => font.baseFont);
    assert.equal(names.length, 3);
    assert.equal(new Set(names).size, 3, `three fonts, three names (${names.join(', ')})`);
    if (fsType === 0) assert.ok(names.every((name) => /^[A-Z]{6}\+BymTest-Regular$/.test(name)), 'subsets keep the name of their font behind their tag');
    else assert.ok(names.indexOf('BymTest-Regular') >= 0, 'one whole font keeps the name as it is');
  }
  assert.ok(first.length > 0 && second.length > 0);
});

// ---------------------------------------------------------------------------
// B7: colours and reasons

test('review: a colour CSS does not take is not read as one', () => {
  for (const wrong of ['rgba(1,2,3/0.5)', 'rgb(1, 2 3)', 'rgb(1 2, 3)', 'rgb(1 2 3, 0.5)', 'rgb(1,2,3 / 0.5)', 'rgb(1,2)', 'rgb(1,2,3,)', 'rgb(1 2 3 /)', 'rgb(1,,2,3)', 'rgb(-1,2,3)', 'rgb(1e2,2,3)', 'rgb(1 2 3 0.5)', 'rgb 1 2 3', '#12345', '#1234567', '#gggggg'])
    assert.equal(parseCssColor(wrong), undefined, wrong);
  assert.deepEqual(parseCssColor('rgb(1, 2, 3)'), { r: 1, g: 2, b: 3, a: 1 });
  assert.deepEqual(parseCssColor('rgba(1,2,3,.5)'), { r: 1, g: 2, b: 3, a: 0.5 });
  assert.deepEqual(parseCssColor('rgb(1 2 3)'), { r: 1, g: 2, b: 3, a: 1 });
  assert.deepEqual(parseCssColor('rgb(1 2 3 / 0.25)'), { r: 1, g: 2, b: 3, a: 0.25 });
  assert.deepEqual(parseCssColor('rgb(1 2 3 / 50%)'), { r: 1, g: 2, b: 3, a: 0.5 });
  assert.deepEqual(parseCssColor('rgba(300, 2, 3, 7)'), { r: 255, g: 2, b: 3, a: 1 }, 'numbers beyond the range are the ends of the range');
  // A page painted with such a colour is not written with a guess.
  const made = createPdfCanvas({ width: 10, height: 10, pageWidth: 10, pageHeight: 10, fonts: new PdfFontSet([]), measure: { width: () => 0 } });
  made.context.fillStyle = 'rgb(1, 2 3)';
  made.context.fillRect(0, 0, 5, 5);
  assert.match(String(made.canvas.unsupported), /colour that is not read/);
});

test('review: a reason is plain text of a bounded length, whatever a file named its font', () => {
  assert.equal(cleanReason('a\u0000b\u001b[31mc\nd\u202ee\u2066f\u0085g'), 'ab[31mc defg');
  assert.equal(cleanReason('x'.repeat(5000)).length, 300);
  assert.equal(cleanReason(undefined as unknown as string), '');
  // The name of a family reaches the reason of a page through the list of fonts of its text.
  const set = new PdfFontSet([], () => true);
  const made = createPdfCanvas({ width: 10, height: 10, pageWidth: 10, pageHeight: 10, fonts: set, measure: { width: () => 0 } });
  made.context.font = '12px "Evil\u0007\u202e Name' + 'x'.repeat(2000) + '"';
  made.context.fillText('a', 0, 0);
  const reason = String(made.canvas.unsupported);
  assert.match(reason, /font of the device \("Evil Name/);
  assert.ok(reason.length <= 300 && !/[\u0000-\u001f\u007f-\u009f\u202a-\u202e\u2066-\u2069]/.test(reason), 'no control characters, and no more than 300 characters');
  const thrown = drawPage({ fonts: set, measure: { width: () => 0 } }, { pixelWidth: 10, pixelHeight: 10, pageWidth: 10, pageHeight: 10 }, () => {
    throw new Error('bad\u0000\u202e' + 'y'.repeat(1000));
  });
  assert.ok(String(thrown.reason).length <= 300 && String(thrown.reason).indexOf('\u0000') < 0 && String(thrown.reason).indexOf('\u202e') < 0);
});

// ---------------------------------------------------------------------------
// B5: pictures of a page that was not drawn

/** A document whose canvases give back what a source says its pixels are, and count how often they are read. */
function pictureDocument(): { doc: Document; reads: () => number } {
  let reads = 0;
  const doc = {
    createElement: () => {
      let drawn: { pixels?: number[] } = {};
      return {
        width: 0,
        height: 0,
        getContext: () => ({
          clearRect: () => undefined,
          drawImage: (source: { pixels?: number[] }) => {
            drawn = source;
          },
          getImageData: (_x: number, _y: number, width: number, height: number) => {
            reads++;
            return { data: drawn.pixels || new Array(width * height * 4).fill(255) };
          },
        }),
      };
    },
  } as unknown as Document;
  return { doc, reads: () => reads };
}

test('review: the pixels of a picture read for a page that was not drawn are let go of; those of a drawn page are kept until written', async () => {
  const { doc, reads } = pictureDocument();
  const images = canvasPictures(doc, 'lossless');
  const set = new PdfFontSet([], () => true);
  const tools = { fonts: set, measure: { width: () => 10 }, images };
  const size = { pixelWidth: 100, pixelHeight: 100, pageWidth: 100, pageHeight: 100 };
  const logo = { width: 1, height: 1, pixels: [1, 2, 3, 255] };
  const photo = { width: 1, height: 1, pixels: [9, 8, 7, 255] };
  // A page that is drawn: its picture is kept, read once.
  const first = drawPage(tools, size, (ctx) => ctx.drawImage(logo as unknown as CanvasImageSource, 0, 0));
  assert.ok(first.vector, 'the first page is drawn');
  const kept = (first.vector as { pictures: PdfPicture[] }).pictures[0];
  // A page that is not drawn (text in a font of the device), with the picture of the first page and one of its own.
  const refused = drawPage(tools, size, (ctx) => {
    ctx.drawImage(logo as unknown as CanvasImageSource, 0, 0);
    ctx.drawImage(photo as unknown as CanvasImageSource, 0, 0);
    ctx.font = '12px Arial';
    ctx.fillText('x', 0, 0);
  });
  assert.equal(refused.vector, undefined);
  assert.equal(reads(), 2, 'each picture was read once so far');
  assert.equal(images.read(logo), kept, 'the picture a drawn page uses is still the same picture');
  assert.equal(reads(), 2);
  // The picture only the refused page asked for was let go of: asked for again, it is read again.
  const again = images.read(photo);
  assert.ok(again && again.image, 'it can be read anew');
  assert.equal(reads(), 3, 'its pixels were not held meanwhile');
  // A picture whose page was written into a file holds no pixels any more, whatever happens to later pages.
  const pdf = new PdfDocument({ date: DATE });
  await pdf.addPage({ width: 100, height: 100, vector: first.vector });
  assert.equal(kept.image, undefined);
});

// ---------------------------------------------------------------------------
// B3: ligatures

/** The ligatures of the font of the package, by the letters they stand for: what a browser draws by default, and for Turkish text (no "fi": it would hide whether the i has its dot). */
const LIGATURES_DEFAULT: { [letters: string]: number } = { ffi: 124, ffl: 125, ff: 123, fi: 126, fl: 127 };
const LIGATURES_TURKISH: { [letters: string]: number } = { ffl: 125, ff: 123, fl: 127 };

/**
 * A browser that shapes text: ligatures by the letters given (longest
 * first), none across a zero width non-joiner, and kerning between the
 * pairs given (px per font size). It shares nothing with the reader of GSUB.
 */
function shapingBrowser(set: PdfFontSet, ligatures: { [letters: string]: number }, kern: (before: number, after: number) => number = () => 0): PdfCanvasMeasure & { glyphs(font: string, text: string): number[] } {
  const keys = Object.keys(ligatures).sort((a, b) => b.length - a.length);
  const shape = (font: string, text: string): { face: PdfFontFace; size: number; glyphs: number[] } => {
    const resolved = set.resolve(font) as { face: PdfFontFace; size: number };
    const glyphs: number[] = [];
    for (const piece of text.split('\u200c')) {
      for (let i = 0; i < piece.length; ) {
        const key = keys.filter((letters) => piece.substr(i, letters.length) === letters)[0];
        if (key) {
          glyphs.push(ligatures[key]);
          i += key.length;
        } else {
          const code = piece.codePointAt(i) as number;
          glyphs.push(resolved.face.glyph(code));
          i += code > 0xffff ? 2 : 1;
        }
      }
      // What stands for the joiner: nothing to see, and nothing kerning reaches across.
      glyphs.push(-1);
    }
    glyphs.pop();
    return { face: resolved.face, size: resolved.size, glyphs };
  };
  const width = (font: string, text: string, kerned: boolean): number => {
    const shaped = shape(font, text);
    let units = 0;
    shaped.glyphs.forEach((glyph, index) => {
      if (glyph < 0) return;
      units += shaped.face.advance(glyph);
      if (kerned && index > 0 && shaped.glyphs[index - 1] >= 0) units += kern(shaped.glyphs[index - 1], glyph) * shaped.face.unitsPerEm;
    });
    return (units * shaped.size) / shaped.face.unitsPerEm;
  };
  return {
    width: (font, text) => width(font, text, true),
    plain: (font, text) => {
      const shaped = shape(font, text);
      // What is drawn: these glyphs, one after the other. Other glyphs, another picture; a joiner draws nothing.
      return { width: width(font, text, false), ink: shaped.glyphs.filter((glyph) => glyph >= 0).join(' ') };
    },
    glyphs: (font, text) => shape(font, text).glyphs.filter((glyph) => glyph >= 0),
  };
}

const BONUM_FACES = { regular: bonum('regular'), bold: bonum('bold'), italic: bonum('italic'), bolditalic: bonum('bolditalic') };
function bonumSet(): PdfFontSet {
  return new PdfFontSet(
    [
      { family: 'Bonum', weight: 400, style: 'normal', bytes: BONUM_FACES.regular },
      { family: 'Bonum', weight: 700, style: 'normal', bytes: BONUM_FACES.bold },
      { family: 'Bonum', weight: 400, style: 'italic', bytes: BONUM_FACES.italic },
      { family: 'Bonum', weight: 700, style: 'italic', bytes: BONUM_FACES.bolditalic },
    ],
    () => false,
  );
}
const FONT_OF: { [face: string]: string } = { regular: '20px Bonum', bold: 'bold 20px Bonum', italic: 'italic 20px Bonum', bolditalic: 'italic bold 20px Bonum' };

/** Writes the words under each other with the measurer of a browser; the page, read back. */
async function wordsPage(measure: PdfCanvasMeasure, font: string, words: string[], set = bonumSet(), spacing?: string): Promise<{ reason?: string; file?: Uint8Array; glyphs: number[][]; texts: string[]; advances: number[] }> {
  const made = createPdfCanvas({ width: 300, height: 30 * words.length + 20, pageWidth: 300, pageHeight: 30 * words.length + 20, fonts: set, measure });
  made.context.font = font;
  if (spacing) (made.context as unknown as { letterSpacing: string }).letterSpacing = spacing;
  words.forEach((word, index) => made.context.fillText(word, 10, 30 + 30 * index));
  if (made.canvas.unsupported !== undefined) return { reason: made.canvas.unsupported, glyphs: [], texts: [], advances: [] };
  const file = await buildPdf([{ width: 300, height: 30 * words.length + 20, vector: made.canvas.finish() }], { date: DATE });
  const page = checkPdf(file).pages[0];
  return { file, glyphs: page.texts.map((text) => text.glyphs), texts: page.texts.map((text) => text.text), advances: page.texts.map((text) => text.advance) };
}

const WORDS = ['office', 'fiyat', 'profil', 'waffle', 'shuffle', 'fjord'];
// What HarfBuzz gives for the words in the four faces of the font (the glyphs have the same numbers in all four):
// hb-shape --no-glyph-names --features=-kern [--language=tr] texgyrebonum-*.otf "office fiyat profil waffle shuffle fjord"
const SHAPED_DEFAULT = [[82, 124, 44, 51], [126, 119, 29, 106], [85, 97, 82, 126, 73], [115, 29, 125, 51], [99, 64, 110, 125, 51], [56, 69, 82, 97, 48]];
const SHAPED_TURKISH = [[82, 123, 67, 44, 51], [56, 67, 119, 29, 106], [85, 97, 82, 56, 67, 73], [115, 29, 125, 51], [99, 64, 110, 125, 51], [56, 69, 82, 97, 48]];
// The advances of those glyphs in the regular face, in thousandths of the font size, summed per word.
const ADVANCES_DEFAULT_REGULAR = [560 + 1000 + 520 + 520, 620 + 540 + 580 + 380, 620 + 440 + 560 + 620 + 300, 780 + 580 + 1000 + 520, 520 + 660 + 680 + 1000 + 520, 320 + 300 + 560 + 440 + 620];

test('review: the ligatures of a font are read from its GSUB table: the lookups of every language of Latin text', () => {
  for (const which of ['regular', 'bold', 'italic', 'bolditalic'] as const) {
    const face = readFont(BONUM_FACES[which]);
    const ligatures = fontLigatures(face);
    assert.equal(ligatures.state, 'known', which);
    assert.equal(ligatures.otherScripts, true, `${which}: the font has rules for Cyrillic too`);
    assert.ok(ligatures.variants.length >= 2, `${which}: more than one language system (${ligatures.variants.length})`);
    const glyphsOf = (text: string): number[] => Array.from(text).map((char) => face.glyph(char.codePointAt(0) as number));
    // One of the language systems shapes as a browser does by default, one as for Turkish.
    const shapes = ligatures.variants.map((variant) => WORDS.map((word) => applyLigatures(variant, glyphsOf(word)).map((entry) => entry.glyph)));
    assert.ok(shapes.some((shape) => JSON.stringify(shape) === JSON.stringify(SHAPED_DEFAULT)), `${which}: a language system that gives what HarfBuzz gives`);
    assert.ok(shapes.some((shape) => JSON.stringify(shape) === JSON.stringify(SHAPED_TURKISH)), `${which}: a language system that gives what HarfBuzz gives for Turkish`);
    assert.deepEqual(
      applyLigatures(ligatures.variants[shapes.findIndex((shape) => JSON.stringify(shape) === JSON.stringify(SHAPED_DEFAULT))], glyphsOf('office')).map((entry) => entry.count),
      [1, 3, 1, 1],
      'a ligature says how many letters it stands for',
    );
  }
});

test('review: a GSUB table of the fonts made here: ligatures, through an extension too, in the order of the lookups', () => {
  // "ABC" is glyph 5, "AB" glyph 9, "BA" glyph 8 (in a second lookup, behind an extension).
  const lookups = [
    { type: 4, ligatures: [[5, 2, 3, 4], [9, 2, 3]] },
    { type: 4, extension: true, ligatures: [[8, 3, 2]] },
  ];
  const face = readFont(testFont({ extra: { GSUB: gsubTable({ lookups, features: [{ tag: 'liga', lookups: [1, 0] }], scripts: [{ tag: 'DFLT', features: [0] }, { tag: 'latn', features: [0], languages: [{ tag: 'TRK ', features: [] }] }] }) } }));
  const ligatures = fontLigatures(face);
  assert.equal(ligatures.state, 'known');
  assert.equal(ligatures.otherScripts, false);
  assert.equal(ligatures.variants.length, 2, 'the lookups of the default language, and none for Turkish');
  const full = ligatures.variants.filter((variant) => variant.length === 2)[0];
  // hb-shape gives [gid5|gid1|gid9|gid1|gid8|gid1|gid2] for "ABC AB BA A".
  assert.deepEqual(
    applyLigatures(full, [2, 3, 4, 1, 2, 3, 1, 3, 2, 1, 2]).map((entry) => entry.glyph),
    [5, 1, 9, 1, 8, 1, 2],
  );
  // "ABA": the first lookup takes "AB", and nothing is left for the second.
  assert.deepEqual(applyLigatures(full, [2, 3, 2]), [{ glyph: 9, count: 2 }, { glyph: 2, count: 1 }]);
  assert.deepEqual(fontLigatures(readFont(testFont())), { state: 'none', variants: [], otherScripts: false }, 'a font without a GSUB table');
  const other = readFont(testFont({ extra: { GSUB: gsubTable({ lookups: [{ type: 1 }], features: [{ tag: 'smcp', lookups: [0] }] }) } }));
  assert.equal(fontLigatures(other).state, 'none', 'features that are off by default are not asked');
});

test('review: a font that replaces glyphs by rules that are not read says so', () => {
  const state = (gsub: number[]): string => fontLigatures(readFont(testFont({ extra: { GSUB: gsub } }))).state;
  const ligature = { type: 4, ligatures: [[9, 2, 3]] };
  assert.equal(state(gsubTable({ lookups: [ligature], features: [{ tag: 'liga', lookups: [0] }] })), 'known');
  assert.equal(state(gsubTable({ lookups: [{ type: 6 }], features: [{ tag: 'liga', lookups: [0] }] })), 'unknown', 'ligatures by context');
  assert.equal(state(gsubTable({ lookups: [{ type: 6, extension: true }], features: [{ tag: 'rlig', lookups: [0] }] })), 'unknown', 'the same behind an extension');
  assert.equal(state(gsubTable({ lookups: [ligature, { type: 1 }], features: [{ tag: 'liga', lookups: [0] }, { tag: 'calt', lookups: [1] }] })), 'unknown', 'alternates by context');
  assert.equal(state(gsubTable({ lookups: [ligature], features: [{ tag: 'liga', lookups: [0] }], scripts: [{ tag: 'cyrl', features: [0] }] })), 'none', 'ligatures of another script only');
  // A table cut short, or one that points outside itself.
  const whole = gsubTable({ lookups: [ligature], features: [{ tag: 'liga', lookups: [0] }] });
  assert.equal(state(whole.slice(0, whole.length - 6)), 'unknown');
  assert.equal(state([0, 1, 0, 0, 0xff, 0xff, 0, 10, 0, 10]), 'unknown');
  // Broken at random it is read or it is unknown, and never anything else.
  const random = seeded(99);
  const started = Date.now();
  for (let round = 0; round < 400; round++) {
    const broken = (round % 2 === 0 ? whole : Array.from(BONUM_FACES.regular.subarray(0, 0))).slice();
    const source = round % 2 === 0 ? broken : undefined;
    let face: PdfFontFace;
    if (source) {
      for (let i = 0; i < 3; i++) source[Math.floor(random() * source.length)] = Math.floor(random() * 256);
      face = readFont(testFont({ extra: { GSUB: source } }));
    } else {
      // The GSUB table of the font of the package, a few bytes changed.
      const bytes = BONUM_FACES.regular.slice();
      const original = readFont(BONUM_FACES.regular).tables.get('GSUB') as { offset: number; length: number };
      for (let i = 0; i < 3; i++) bytes[original.offset + Math.floor(random() * (round % 4 === 1 ? 64 : original.length))] = Math.floor(random() * 256);
      face = readFont(bytes);
    }
    const read = fontLigatures(face);
    assert.ok(read.state === 'none' || read.state === 'known' || read.state === 'unknown', `round ${round}`);
    if (read.state === 'known') for (const variant of read.variants) applyLigatures(variant, [56, 56, 67, 2, 3, 4]);
  }
  assert.ok(Date.now() - started < 20000, 'no broken table takes long');
});

popplerTest('review: words with ligatures are written with the glyphs the browser draws, and read back letter by letter', async () => {
  for (const which of ['regular', 'bold', 'italic', 'bolditalic'] as const) {
    const set = bonumSet();
    const page = await wordsPage(shapingBrowser(set, LIGATURES_DEFAULT), FONT_OF[which], WORDS, set);
    assert.equal(page.reason, undefined, which);
    assert.deepEqual(page.glyphs, SHAPED_DEFAULT, `${which}: the glyphs HarfBuzz gives`);
    assert.deepEqual(page.texts, WORDS, `${which}: each glyph of a ligature stands for its letters`);
    const report = poppler(page.file as Uint8Array);
    assert.equal(report.complaints, '', which);
    assert.deepEqual(report.text[0].trim().split('\n'), WORDS, `${which}: Poppler reads the letters`);
    if (which === 'regular') {
      const words = popplerWords(page.file as Uint8Array).words;
      words.forEach((word, index) => {
        assert.equal(word.text, WORDS[index]);
        assert.ok(Math.abs(word.xMin - 10) < 0.05 && Math.abs(word.xMax - (10 + ADVANCES_DEFAULT_REGULAR[index] * 0.02)) < 0.05, `"${word.text}" is as wide as its glyphs with the ligature (${word.xMin} to ${word.xMax})`);
      });
    }
  }
});

test('review: where the browser draws other ligatures (Turkish text), or none, those are written', async () => {
  for (const which of ['regular', 'bold', 'italic', 'bolditalic'] as const) {
    const set = bonumSet();
    const turkish = await wordsPage(shapingBrowser(set, LIGATURES_TURKISH), FONT_OF[which], WORDS, set);
    assert.equal(turkish.reason, undefined, which);
    assert.deepEqual(turkish.glyphs, SHAPED_TURKISH, `${which}: what HarfBuzz gives for Turkish`);
    assert.deepEqual(turkish.texts, WORDS, which);
    // A browser that draws no ligatures at all.
    const none = shapingBrowser(set, {});
    const plain = await wordsPage(none, FONT_OF[which], WORDS, set);
    assert.deepEqual(plain.glyphs, WORDS.map((word) => none.glyphs(FONT_OF[which], word)), `${which}: every letter its own glyph`);
    assert.ok(plain.glyphs[0].length === 6, 'six glyphs for "office"');
  }
  // Kerning beside a ligature moves the ligature, and is no reason to doubt it.
  const set = bonumSet();
  const kerned = await wordsPage(shapingBrowser(set, LIGATURES_DEFAULT, (before, after) => (before === 82 && after === 124 ? -0.03 : before === 124 ? 0.02 : 0)), FONT_OF.regular, ['office'], set);
  assert.deepEqual(kerned.glyphs, [SHAPED_DEFAULT[0]]);
  assert.ok(Math.abs(kerned.advances[0] - (2600 * 0.02 - 0.6 + 0.4)) < 0.01, `the word is as wide as the browser measured it (${kerned.advances[0]})`);
});

test('review: spaced letters are not joined, as a browser does not join them', async () => {
  const set = bonumSet();
  const browser = shapingBrowser(set, LIGATURES_DEFAULT);
  // With letter spacing a browser draws no ligatures: its widths are those of the letters.
  const spaced: PdfCanvasMeasure = { width: (font, text) => shapingBrowser(set, {}).width(font, text), plain: browser.plain };
  const page = await wordsPage(spaced, FONT_OF.regular, ['office'], set, '1px');
  assert.equal(page.reason, undefined);
  assert.deepEqual(page.glyphs, [[82, 56, 56, 67, 44, 51]]);
});

test('review: where it cannot be told which ligatures the browser drew, the page is not written', async () => {
  const set = bonumSet();
  const browser = shapingBrowser(set, LIGATURES_DEFAULT);
  // A browser that cannot measure without kerning, and gives no ink: nothing tells "ffi" from "ff" and "i".
  const blind = await wordsPage({ width: browser.width }, FONT_OF.regular, ['office'], set);
  assert.match(String(blind.reason), /ligature/);
  // Words without letters a ligature starts with need no such measuring.
  const plainWords = await wordsPage({ width: browser.width }, FONT_OF.regular, ['kalem', 'masa'], set);
  assert.equal(plainWords.reason, undefined);
  // A browser whose measures fit none of what the font could give.
  const odd = await wordsPage({ width: browser.width, plain: (font, text) => ({ width: browser.width(font, text) + (text.indexOf('\u200c') >= 0 ? 0 : 3), ink: text }) }, FONT_OF.regular, ['office'], set);
  assert.match(String(odd.reason), /ligature/);
  // A ligature that is as wide as its letters, in a browser that may or may not draw it: only the picture tells.
  // (In the regular face "fi" is 620 wide, and so are "f" and "i".) The browser draws the ligature: its picture
  // of "fi" is not that of the two letters held apart.
  const drawn = await wordsPage(browser, FONT_OF.regular, ['fi'], set);
  assert.deepEqual(drawn.glyphs, [[126]]);
  // The browser draws the letters (Turkish text): the two pictures are the same, and the letters are written.
  const letters = await wordsPage(shapingBrowser(set, LIGATURES_TURKISH), FONT_OF.regular, ['fi'], set);
  assert.deepEqual(letters.glyphs, [[56, 67]]);
  // A browser that says the same of both whatever it draws leaves nothing to tell them by... the letters are what
  // looks the same as what it drew.
  const same = await wordsPage({ width: browser.width, plain: (font, text) => ({ width: (browser.plain as NonNullable<typeof browser.plain>)(font, text)!.width, ink: 'x' }) }, FONT_OF.regular, ['fi'], set);
  assert.deepEqual(same.glyphs, [[56, 67]]);
  // A font with rules that are not read: none of its text is written.
  const unknown = testFont({ extra: { GSUB: gsubTable({ lookups: [{ type: 6 }], features: [{ tag: 'liga', lookups: [0] }] }) } });
  const unread = new PdfFontSet([{ family: 'Bonum', weight: 400, style: 'normal', bytes: unknown }], () => false);
  assert.match(String((await wordsPage(shapingBrowser(unread, {}), '20px Bonum', ['AB'], unread)).reason), /replaces glyphs by rules that are not read/);
  // Text of a script whose rules are not read, in a font that has such rules; without such rules it is written.
  const ruled = testFont({ map: { 0x41f: 2, 0x440: 3 }, extra: { GSUB: gsubTable({ lookups: [{ type: 4, ligatures: [[9, 2, 3]] }], features: [{ tag: 'liga', lookups: [0] }], scripts: [{ tag: 'DFLT', features: [] }, { tag: 'cyrl', features: [0] }] }) } });
  const ruledSet = new PdfFontSet([{ family: 'Bonum', weight: 400, style: 'normal', bytes: ruled }], () => false);
  const cyrillic = String.fromCharCode(0x41f, 0x440);
  assert.match(String((await wordsPage(shapingBrowser(ruledSet, {}), '20px Bonum', [cyrillic], ruledSet)).reason), /script the font BymTest-Regular has rules for/);
  const unruled = testFont({ map: { 0x41f: 2, 0x440: 3 } });
  const unruledSet = new PdfFontSet([{ family: 'Bonum', weight: 400, style: 'normal', bytes: unruled }], () => false);
  assert.deepEqual((await wordsPage(shapingBrowser(unruledSet, {}), '20px Bonum', [cyrillic], unruledSet)).glyphs, [[2, 3]]);
});

test('review: a ligature of the fonts made here is written when the browser draws it, in the map to Unicode with all its letters', async () => {
  // "AB" is glyph 9: 900 wide where "A" and "B" are 1200.
  const bytes = testFont({ extra: { GSUB: gsubTable({ lookups: [{ type: 4, ligatures: [[9, 2, 3]] }], features: [{ tag: 'liga', lookups: [0] }] }) } });
  const set = new PdfFontSet([{ family: 'Bonum', weight: 400, style: 'normal', bytes }], () => false);
  const joined = await wordsPage(shapingBrowser(set, { AB: 9 }), '20px Bonum', ['ABAB C', 'BA'], set);
  assert.deepEqual(joined.glyphs, [[9, 9, 1, 4], [3, 2]]);
  assert.deepEqual(joined.texts, ['ABAB C', 'BA']);
  assert.ok(Math.abs(joined.advances[0] - (900 + 900 + 300 + 400) * 0.02) < 0.001);
  const checked = checkPdf(joined.file as Uint8Array);
  assert.deepEqual(Array.from(checked.fonts[0].widths.keys()), [1, 2, 3, 4, 9]);
  const apart = await wordsPage(shapingBrowser(set, {}), '20px Bonum', ['ABAB C'], set);
  assert.deepEqual(apart.glyphs, [[2, 3, 2, 3, 1, 4]]);
});

// ---------------------------------------------------------------------------
// B4: long runs

test('review: in a run of more glyphs than are placed one by one every word still stands where the browser has it', async () => {
  const set = bonumSet();
  // Kerning as a browser applies it: every pair with a capital letter in front stands closer, every pair behind a comma further.
  const kern = (before: number, after: number): number => (before < 60 && after > 60 ? -0.04 : before === 104 ? 0.01 : 0);
  const browser = shapingBrowser(set, {}, kern);
  const sentence = 'Tavla Yaz Vakit Kalem Toprak Yurt Vatan Para Tane Ayva Pazar Waka ';
  let text = '';
  while (Array.from(text).length < 339) text += sentence;
  text = text.slice(0, 339);
  const made = createPdfCanvas({ width: 8000, height: 100, pageWidth: 8000, pageHeight: 100, fonts: set, measure: browser });
  made.context.font = '20px Bonum';
  made.context.fillText(text, 10, 50);
  assert.equal(made.canvas.unsupported, undefined);
  const page = checkPdf(await buildPdf([{ width: 8000, height: 100, vector: made.canvas.finish() }], { date: DATE })).pages[0];
  assert.equal(page.text, text, 'the whole run is written');
  assert.ok(page.texts.length >= 2, 'in more than one piece');
  assert.ok(page.texts.every((piece) => piece.glyphs.length <= 256 && piece.spacing === 0), 'no piece is longer than what is placed glyph by glyph, and none is spaced out');
  // Where every glyph of the run starts, on the page.
  const starts: number[] = [];
  for (const piece of page.texts) for (const start of piece.starts) starts.push(piece.x + start);
  assert.equal(starts.length, 339);
  let worst = 0;
  for (let i = 0; i < text.length; i++) {
    if (i > 0 && text.charAt(i - 1) !== ' ') continue;
    // A word starts where the text before it ends as the browser measures it... less what kerning moves its first letter by.
    const before = i === 0 ? 0 : browser.width('20px Bonum', text.slice(0, i + 1)) - browser.width('20px Bonum', text.charAt(i));
    worst = Math.max(worst, Math.abs(starts[i] - 10 - before));
  }
  assert.ok(worst < 0.02 * 20, `the edge of a word is off by ${(worst / 20).toFixed(4)} em at most`);
});

/** A document whose canvases measure ten px a character (one less with kerning) and paint, as pixels, the codes of what they were asked to draw. */
function inkDocument(options: { kerning?: boolean | 'stubborn' | 'stubborn at first'; pixels?: boolean } = {}): { doc: Document; calls: string[]; contexts: { [name: string]: unknown }[] } {
  const calls: string[] = [];
  const contexts: { [name: string]: unknown }[] = [];
  const doc = {
    createElement: () => {
      const canvas: { width: number; height: number; getContext(): unknown } = {
        width: 300,
        height: 150,
        getContext: () => {
          let drawn = '';
          const context: { [name: string]: unknown } = {
            font: '',
            textBaseline: 'alphabetic',
            fillStyle: '',
            measureText(text: string) {
              calls.push(`measure|${String(this.fontKerning)}|${String(this.font)}|${text}`);
              return { width: text.length * 10 - (this.fontKerning === 'none' ? 0 : 1), actualBoundingBoxAscent: 7 };
            },
            clearRect: () => {
              drawn = '';
            },
            fillText(text: string, x: number, y: number) {
              calls.push(`draw|${String(this.fontKerning)}|${String(this.font)}|${text}`);
              drawn += `${text}@${x},${y}`;
            },
            getImageData: (_x: number, _y: number, width: number, height: number) => {
              if (options.pixels === false) throw new Error('no pixels here');
              // The picture of a text: its letters without the joiners, which draw nothing.
              const data = new Uint8Array(width * height * 4);
              const seen = drawn.replace(/\u200c/g, '');
              for (let i = 0; i < seen.length; i++) data[i * 4 + 3] = seen.charCodeAt(i) & 255;
              return { data };
            },
          };
          if (options.kerning === 'stubborn' || (options.kerning === 'stubborn at first' && contexts.length === 0)) Object.defineProperty(context, 'fontKerning', { get: () => 'auto', set: () => undefined });
          else if (options.kerning !== false) context.fontKerning = 'auto';
          contexts.push(context);
          return context;
        },
      };
      return canvas;
    },
  } as unknown as Document;
  return { doc, calls, contexts };
}

test('review: the real canvas is asked without kerning for the width of a text and for its picture, and left as it was', () => {
  const page = inkDocument();
  const measurer = canvasMeasure(page.doc);
  assert.ok(measurer && measurer.plain, 'a document with canvases can be asked');
  const joined = measurer.plain('italic 12px "A"', 'ffi');
  assert.ok(joined, 'a canvas that can turn kerning off says what it draws');
  assert.equal(joined.width, 30, 'the width without kerning, at the size of the font');
  assert.equal(measurer.width('italic 12px "A"', 'ffi'), 29, 'widths are measured with kerning as before');
  assert.ok(page.contexts.every((context) => context.fontKerning === 'auto'), 'kerning is on again afterwards');
  // The picture is painted larger than the text is, so that small differences are pixels.
  assert.ok(page.calls.some((call) => /^draw\|none\|italic 64px "A"\|ffi$/.test(call)), `the text is painted at 64 px without kerning (${page.calls.join('; ')})`);
  // Letters held apart by a joiner look the same when the canvas joins nothing, and the picture says so.
  const apart = measurer.plain('italic 12px "A"', 'ff\u200ci');
  assert.ok(apart && apart.ink === joined.ink, 'the same picture is the same ink');
  const other = measurer.plain('italic 12px "A"', 'ffl');
  assert.ok(other && other.ink !== joined.ink, 'another picture is another ink');
  const before = page.calls.length;
  assert.deepEqual(measurer.plain('italic 12px "A"', 'ffi'), joined);
  assert.equal(page.calls.length, before, 'a text is asked about once');
  // A canvas that cannot turn kerning off, one that does not keep the setting, and one whose pixels cannot be read, cannot be asked.
  for (const options of [{ kerning: false }, { kerning: 'stubborn' as const }, { kerning: 'stubborn at first' as const }, { pixels: false }]) {
    const none = canvasMeasure(inkDocument(options).doc);
    assert.equal((none as { plain: (font: string, text: string) => unknown }).plain('12px A', 'ffi'), undefined, JSON.stringify(options));
  }
  // A text too long to paint is not asked about.
  assert.equal(measurer.plain('12px A', 'f'.repeat(500)), undefined);
});

// ---------------------------------------------------------------------------
// B8: what changing the code by a hair did not make any test notice

interface Painted {
  ctx: CanvasRenderingContext2D;
  unsupported(): string | undefined;
  page(): Promise<ReturnType<typeof checkPdf>['pages'][0]>;
  checked(): Promise<ReturnType<typeof checkPdf>>;
}

function painted(options: { width?: number; height?: number; pageWidth?: number; pageHeight?: number; fonts?: PdfFontSet; measure?: PdfCanvasMeasure; pictures?: Map<unknown, PdfPicture> } = {}): Painted {
  const fonts = options.fonts || new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes: TEST_FONT }], () => false);
  const pageWidth = options.pageWidth || 200;
  const pageHeight = options.pageHeight || 100;
  const pictures = options.pictures;
  const measure: PdfCanvasMeasure = options.measure || {
    width: (font, text) => {
      const resolved = fonts.resolve(font) as { face: PdfFontFace; size: number };
      let units = 0;
      for (const char of text) units += resolved.face.advance(resolved.face.glyph(char.codePointAt(0) as number));
      return (units * resolved.size) / resolved.face.unitsPerEm;
    },
  };
  const made = createPdfCanvas({ width: options.width || pageWidth, height: options.height || pageHeight, pageWidth, pageHeight, fonts, measure, images: pictures ? { read: (source) => pictures.get(source) } : undefined });
  const checked = async () => checkPdf(await buildPdf([{ width: pageWidth, height: pageHeight, vector: made.canvas.finish() }], { date: DATE }));
  return { ctx: made.context, unsupported: () => made.canvas.unsupported, checked, page: async () => (await checked()).pages[0] };
}

const TEST_FONT = testFont();
const close = (actual: (string | number)[], expected: number[], what: string): void => {
  const numbers = actual.filter((entry) => typeof entry === 'number') as number[];
  assert.equal(numbers.length, expected.length, `${what}: ${JSON.stringify(actual)}`);
  for (let i = 0; i < expected.length; i++) assert.ok(Math.abs(numbers[i] - expected[i]) <= 0.002, `${what}: ${JSON.stringify(numbers)} is not ${JSON.stringify(expected)}`);
};

test('review: transforms are applied in the order they were given: moved, then turned, then stretched', async () => {
  const made = painted();
  const ctx = made.ctx;
  ctx.translate(100, 20);
  ctx.rotate(Math.PI / 2);
  ctx.scale(2, 0.5);
  // Stretched along what is now the way down the page, flattened along what is now the way to the left.
  ctx.fillRect(5, 10, 10, 20);
  const path = (await made.page()).shapes[0].path;
  // A point (x, y) lies at (100 - 0.5 y, 20 + 2 x) of the canvas, and at 100 less that from the bottom of the page.
  close(path[0], [95, 70], 'the first corner');
  close(path[1], [95, 50], 'along the x axis of the canvas, turned downwards');
  close(path[2], [85, 50], 'the opposite corner');
  close(path[3], [85, 70], 'the last corner');
  // The matrix a painter reads back is the one that is used.
  const read = (ctx as unknown as { getTransform(): { a: number; b: number; c: number; d: number; e: number; f: number } }).getTransform();
  close([read.a, read.b, read.c, read.d, read.e, read.f], [0, 2, -0.5, 0, 100, 20], 'the matrix of the canvas');
  // transform() multiplies from the right too: a matrix that neither commutes with a turn nor with a stretch.
  const sheared = painted();
  sheared.ctx.scale(2, 1);
  sheared.ctx.transform(1, 0, 1, 1, 10, 0);
  sheared.ctx.fillRect(0, 0, 10, 10);
  const skew = (await sheared.page()).shapes[0].path;
  close(skew[0], [20, 100], 'a point is sheared, then stretched');
  close(skew[2], [60, 90], 'the far corner: (10 + 10 + 10) twice across, 10 down');
});

test('review: a canvas whose pixels are not square on the page keeps its two scales apart', async () => {
  // 200 x 400 pixels for a page of 200 x 100 points: a pixel is one point wide and a quarter high.
  const made = painted({ width: 200, height: 400 });
  made.ctx.fillRect(20, 40, 60, 80);
  made.ctx.beginPath();
  made.ctx.moveTo(10, 400);
  made.ctx.lineTo(200, 0);
  made.ctx.closePath();
  made.ctx.fill();
  const shapes = (await made.page()).shapes;
  assert.deepEqual(shapes[0].path, [['re', 20, 70, 60, 20]]);
  assert.deepEqual(shapes[1].path.slice(0, 2), [['m', 10, 0], ['l', 200, 100]]);
});

test('review: arcs go the way they are asked to, by as much as they are asked to', async () => {
  const reach = (4 / 3) * Math.tan(Math.PI / 8) * 30;
  // Clockwise from the left to the bottom is three quarters of a turn, not a quarter back.
  const three = painted();
  three.ctx.beginPath();
  three.ctx.arc(100, 50, 30, Math.PI, Math.PI / 2);
  three.ctx.stroke();
  const path = (await three.page()).shapes[0].path;
  assert.equal(path.length, 4, 'a start and three quarter turns');
  close(path[0], [70, 50], 'it starts at the left');
  close(path[1], [70, 50 + reach, 100 - reach, 80, 100, 80], 'first to the top of the page (the top of the canvas is up on the page)');
  close(path[3].slice(5), [100, 20], 'and ends at the bottom');
  // A full turn the other way round goes the other way round.
  const back = painted();
  back.ctx.beginPath();
  back.ctx.arc(100, 50, 30, 0, -Math.PI * 2, true);
  back.ctx.stroke();
  const turn = (await back.page()).shapes[0].path;
  assert.equal(turn.length, 5, 'a start and four quarter turns');
  close(turn[1], [130, 50 + reach, 100 + reach, 80, 100, 80], 'counter-clockwise on the canvas: first up the page');
  // Counter-clockwise from the right to the bottom is three quarters, through the top.
  const most = painted();
  most.ctx.beginPath();
  most.ctx.arc(100, 50, 30, 0, Math.PI / 2, true);
  most.ctx.stroke();
  const far = (await most.page()).shapes[0].path;
  assert.equal(far.length, 4);
  close(far[3].slice(5), [100, 20], 'it ends at the bottom of the circle');
  close(far[1].slice(5), [100, 80], 'having gone through its top first');
});

test('review: after a path is closed the next piece starts where the closed one did', async () => {
  const made = painted();
  const ctx = made.ctx;
  ctx.beginPath();
  ctx.moveTo(10, 10);
  ctx.lineTo(50, 10);
  ctx.lineTo(50, 50);
  ctx.closePath();
  // A curve from where the path stands: (10, 10) again, not (50, 50).
  ctx.quadraticCurveTo(10, 40, 40, 70);
  ctx.stroke();
  const path = (await made.page()).shapes[0].path;
  assert.deepEqual(path[3], ['h']);
  close(path[4], [10, 70, 20, 50, 40, 30], 'the curve leaves from the start of the closed piece');
});

test('review: a line has the opacity of its colour and of the canvas together, its dashes the scale of the canvas', async () => {
  const made = painted({ width: 400, height: 200 });
  const ctx = made.ctx;
  ctx.setTransform(2, 0, 0, 2, 0, 0);
  ctx.strokeStyle = 'rgba(0, 0, 255, 0.5)';
  ctx.globalAlpha = 0.5;
  ctx.setLineDash([5]);
  ctx.lineDashOffset = 3;
  ctx.strokeRect(10, 10, 50, 50);
  ctx.setLineDash([1, 2, 3]);
  ctx.strokeRect(10, 10, 50, 50);
  // Under a canvas twice as large again the dashes and where they start are twice as long.
  ctx.scale(2, 2);
  ctx.strokeRect(10, 10, 20, 20);
  const page = await made.page();
  assert.deepEqual(page.shapes.map((shape) => shape.alpha), [0.25, 0.25, 0.25]);
  assert.deepEqual(page.shapes[0].dash, [5, 5], 'a list of one length is that length twice');
  assert.deepEqual(page.shapes[1].dash, [1, 2, 3, 1, 2, 3], 'a list of three is the list twice');
  assert.deepEqual(page.shapes[2].dash, [2, 4, 6, 2, 4, 6]);
  assert.match(page.content, /\[5 5\] 3 d/, 'the dashes start three in');
  assert.match(page.content, /\[2 4 6 2 4 6\] 6 d/, 'and six in under the larger canvas');
  assert.deepEqual((made.ctx as unknown as { getLineDash(): number[] }).getLineDash(), [1, 2, 3, 1, 2, 3]);
});

test('review: a clip by the even-odd rule is a clip by that rule', async () => {
  const made = painted();
  const ctx = made.ctx;
  ctx.beginPath();
  ctx.rect(10, 10, 80, 80);
  ctx.rect(30, 30, 40, 40);
  ctx.clip('evenodd');
  ctx.fillRect(0, 0, 200, 100);
  ctx.beginPath();
  ctx.rect(0, 0, 5, 5);
  ctx.clip();
  const content = (await made.page()).content;
  assert.equal((content.match(/W\* n/g) || []).length, 1, 'the first clip leaves its hole out');
  assert.equal((content.match(/W n/g) || []).length, 1, 'the second one is by the other rule');
});

test('review: what a picture needed of the page is forgotten with it: the next thing is painted as the canvas says', async () => {
  const source = {};
  const picture: PdfPicture = { width: 1, height: 1, image: { kind: 'samples', width: 1, height: 1, colors: 1, data: Uint8Array.from([0]) } };
  const made = painted({ pictures: new Map([[source, picture]]) });
  const ctx = made.ctx;
  ctx.globalAlpha = 0.5;
  ctx.fillRect(0, 0, 10, 10);
  // The picture is opaque; so is what follows it.
  ctx.globalAlpha = 1;
  ctx.drawImage(source as CanvasImageSource, 20, 20);
  ctx.fillRect(40, 0, 10, 10);
  ctx.globalAlpha = 0.5;
  ctx.fillRect(60, 0, 10, 10);
  assert.deepEqual((await made.page()).shapes.map((shape) => shape.alpha), [0.5, 1, 0.5]);
});

test('review: a character map that points past its own end gives no glyph', () => {
  // One segment for "A" to "C" whose glyphs are looked up in a list the table has no room for; what lies behind
  // the table of the map (another table of the font) is not the list.
  const tables = testFontTables();
  const map = [0, 4, 0, 32, 0, 0, 0, 4, 0, 4, 0, 1, 0, 0, 0, 0x43, 0xff, 0xff, 0, 0, 0, 0x41, 0xff, 0xff, 0, 0, 0, 1, 0, 4, 0, 0];
  // Behind the map, inside the cmap table: numbers that would be taken for glyphs.
  tables.cmap = [0, 0, 0, 1, 0, 3, 0, 1, 0, 0, 0, 12].concat(map, [0, 3, 0, 3, 0, 3, 0, 3]);
  const face = readFont(sfntFile(tables));
  for (const code of [0x41, 0x42, 0x43]) assert.equal(face.glyph(code), 0, `U+${code.toString(16)} is looked up beyond the map`);
  // With room for the list the same map gives its glyphs.
  const roomy = map.slice();
  roomy[3] = 38;
  tables.cmap = [0, 0, 0, 1, 0, 3, 0, 1, 0, 0, 0, 12].concat(roomy, [0, 2, 0, 3, 0, 4]);
  const sound = readFont(sfntFile(tables));
  assert.deepEqual([0x41, 0x42, 0x43, 0x44].map((code) => sound.glyph(code)), [2, 3, 4, 0]);
});

test('review: the widths of a font are in thousandths of its size, whatever units the font counts in', async () => {
  // The glyphs of the test font twice as large, in a font of 2000 units per em: the same letters.
  const glyphs = TEST_GLYPHS.map((glyph) => ({
    ...glyph,
    advance: glyph.advance * 2,
    contours: glyph.contours ? glyph.contours.map((contour) => contour.map(([x, y]) => [x * 2, y * 2] as [number, number])) : undefined,
    parts: glyph.parts ? glyph.parts.map((part) => ({ ...part, dx: part.dx * 2, dy: part.dy * 2 })) : undefined,
  }));
  const fonts = new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes: testFont({ unitsPerEm: 2000, glyphs }) }], () => false);
  const made = painted({ fonts });
  made.ctx.font = '20px T';
  made.ctx.fillText('AB', 10, 50);
  const checked = await made.checked();
  assert.deepEqual(Array.from(checked.fonts[0].widths.entries()), [[2, 700], [3, 500]]);
  assert.ok(Math.abs(checked.pages[0].texts[0].advance - 24) < 0.001, 'and the text is as wide as on the canvas');
  // The width of a glyph the list does not name is that of the glyph for missing characters, scaled the same way.
  let dw: unknown;
  checked.objects.forEach((object) => {
    const value = object.value as { Subtype?: { name?: string }; DW?: unknown };
    if (value && value.Subtype && value.Subtype.name === 'CIDFontType2') dw = value.DW;
  });
  assert.equal(dw, 500, '/DW');
});

test('review: a glyph two pages use for two characters is copied as the plainer one, whichever page came first', async () => {
  // "A" and the Greek capital alpha are one glyph.
  const bytes = testFont({ map: { 0x20: 1, 0x41: 2, 0x391: 2 } });
  const fonts = new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes }], () => false);
  const page = (text: string) => {
    const made = createPdfCanvas({ width: 200, height: 100, pageWidth: 200, pageHeight: 100, fonts, measure: { width: (_font, value) => value.length * 14 } });
    made.context.font = '20px T';
    made.context.fillText(text, 10, 50);
    return { width: 200, height: 100, vector: made.canvas.finish() };
  };
  for (const order of [['Α', 'A'], ['A', 'Α']]) {
    const checked = checkPdf(await buildPdf(order.map(page), { date: DATE }));
    assert.deepEqual(checked.pages.map((entry) => entry.text), ['A', 'A'], `written ${order.map((text) => text.charCodeAt(0).toString(16)).join(' then ')}`);
  }
});

test('review: a drawn page has no invisible text over it, whatever text is handed along with it', async () => {
  const made = painted();
  made.ctx.font = '20px T';
  made.ctx.fillText('AB', 10, 50);
  const vector = (await (async () => {
    const canvas = createPdfCanvas({ width: 200, height: 100, pageWidth: 200, pageHeight: 100, fonts: new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes: TEST_FONT }], () => false), measure: { width: () => 24 } });
    canvas.context.font = '20px T';
    canvas.context.fillText('AB', 10, 50);
    return canvas.canvas.finish();
  })());
  const checked = checkPdf(await buildPdf([{ width: 200, height: 100, vector, text: { runs: [{ text: 'hidden', x: 10, y: 50, width: 60, size: 20 }], width: 200 }, fill: [255, 0, 0] }], { date: DATE }));
  assert.deepEqual(checked.pages[0].texts.map((text) => [text.text, text.mode]), [['AB', 0]]);
  assert.equal(checked.fonts.length, 1, 'no font for hidden text');
  assert.equal(checked.pages[0].shapes.length, 0, 'and no paper painted under it');
});

test('review: text is placed by its end as by its right, and alignment counts the spacing of letters', async () => {
  const made = painted();
  const ctx = made.ctx;
  ctx.font = '20px T';
  ctx.textAlign = 'end';
  ctx.fillText('A', 100, 50);
  (ctx as unknown as { letterSpacing: string }).letterSpacing = '3px';
  ctx.textAlign = 'center';
  ctx.fillText('AB', 100, 80);
  ctx.textAlign = 'right';
  ctx.fillText('AB', 100, 95);
  const texts = (await made.page()).texts;
  assert.deepEqual(texts.map((text) => text.x), [86, 100 - (24 + 6) / 2, 100 - (24 + 6)]);
});

test('review: two ways to join a text that are painted the same and are equally fine leave nothing to choose by', async () => {
  // "AB" is glyph 10 in one language and "BC" glyph 9 in another; each ligature is as wide as its letters.
  const glyphs = TEST_GLYPHS.concat([{ advance: 1200, contours: [[[0, 0], [0, 10], [10, 10], [10, 0]]] }]);
  const gsub = gsubTable({
    lookups: [{ type: 4, ligatures: [[10, 2, 3]] }, { type: 4, ligatures: [[9, 3, 4]] }],
    features: [{ tag: 'liga', lookups: [0] }, { tag: 'liga', lookups: [1] }],
    scripts: [{ tag: 'latn', features: [0], languages: [{ tag: 'TRK ', features: [1] }] }],
  });
  const set = new PdfFontSet([{ family: 'Bonum', weight: 400, style: 'normal', bytes: testFont({ glyphs, extra: { GSUB: gsub } }) }], () => false);
  const widths = shapingBrowser(set, {});
  // A browser whose picture of the text is that of both ways with one joiner, and of no way with two.
  const joiners = (text: string): number => text.split(String.fromCharCode(0x200c)).length - 1;
  const both = await wordsPage({ width: widths.width, plain: (font, text) => ({ width: widths.width(font, text), ink: joiners(text) <= 1 ? 'one picture' : 'another' }) }, '20px Bonum', ['ABC'], set);
  assert.match(String(both.reason), /cannot be told which ligatures/);
  // Where only one of them is painted like the text, that one is written.
  const second = await wordsPage({ width: widths.width, plain: (font, text) => ({ width: widths.width(font, text), ink: text === 'ABC' || text === 'A' + String.fromCharCode(0x200c) + 'BC' ? 'one picture' : text }) }, '20px Bonum', ['ABC'], set);
  assert.deepEqual(second.glyphs, [[2, 9]]);
});

test('review: the pieces of a long run of spaced letters stand where the spacing before them puts them', async () => {
  const fonts = new PdfFontSet([{ family: 'T', weight: 400, style: 'normal', bytes: TEST_FONT }], () => false);
  const made = painted({ fonts, pageWidth: 9000 });
  made.ctx.font = '20px T';
  (made.ctx as unknown as { letterSpacing: string }).letterSpacing = '2px';
  made.ctx.fillText('AB '.repeat(100), 0, 50);
  const texts = (await made.page()).texts;
  assert.deepEqual(texts.map((text) => text.glyphs.length), [255, 45]);
  assert.ok(texts.every((text) => text.spacing === 2), 'every piece is spaced');
  // 85 times "AB " of 30 px, and 255 letters each followed by 2 px.
  assert.ok(Math.abs(texts[1].x - (85 * 30 + 255 * 2)) < 0.01, `the second piece starts at ${texts[1].x}`);
});
