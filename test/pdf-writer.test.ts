// The core of the PDF writer: values, objects, streams, the cross reference
// table, the trailer; the glyphless font and its ToUnicode map; whole files,
// read back by a reader of our own (test/pdf-fixtures) and by Poppler.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { inflateSync } from 'node:zlib';
import { PdfDocument, buildPdf, type PdfPage } from '../src/pdf/document';
import { adler32, zlibDeflate, zlibOwn } from '../src/pdf/flate';
import { FONT_NAME, GLYPH_WIDTH, MAX_FONT_CODES, PdfTextFont, glyphlessFont } from '../src/pdf/font';
import { imageOfPixels, jpegInfo } from '../src/pdf/image';
import { MAX_PDF_BYTES, PdfError, PdfName, PdfRef, PdfString, PdfWriter, joinParts, latin1, name, pdfDate, pdfNumber, pdfText, serialize } from '../src/pdf/writer';
import { checkPdf, checkTrueType, parseToUnicode, poppler, popplerTest } from './pdf-fixtures';

const text = (bytes: Uint8Array): string => Buffer.from(bytes).toString('latin1');

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof PdfError, `a PdfError is thrown, not ${String(error)}`);
    return error.code;
  }
  return 'no error';
}

async function rejection(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    assert.ok(error instanceof PdfError, `a PdfError is thrown, not ${String(error)}`);
    return error.code;
  }
  return 'no error';
}

/** A file of one empty page, written with the writer alone. */
function minimal(writer: PdfWriter): Uint8Array {
  const pages = writer.reserve();
  const page = writer.add({ Type: name('Page'), Parent: pages, MediaBox: [0, 0, 200, 100], Resources: {} });
  writer.add({ Type: name('Pages'), Kids: [page], Count: 1 }, pages);
  const root = writer.add({ Type: name('Catalog'), Pages: pages });
  writer.finish({ root });
  return joinParts(writer.takeParts());
}

// ------------------------------------------------------------------ numbers

test('numbers are written with digits only, whatever their size', () => {
  assert.equal(pdfNumber(0), '0');
  assert.equal(pdfNumber(-0), '0');
  assert.equal(pdfNumber(12), '12');
  assert.equal(pdfNumber(-7), '-7');
  assert.equal(pdfNumber(0.5), '0.5');
  assert.equal(pdfNumber(595.276), '595.276');
  assert.equal(pdfNumber(1 / 3), '0.3333');
  assert.equal(pdfNumber(2 / 3, 2), '0.67');
  assert.equal(pdfNumber(1e-7), '0', 'a number too small for the digits is 0, not "1e-7"');
  assert.equal(pdfNumber(-1e-7), '0', 'and not "-0"');
  assert.equal(pdfNumber(1.00004), '1');
  assert.equal(pdfNumber(123456789012), '123456789012');
  assert.equal(pdfNumber(1234567.125), '1234567.125');
  assert.doesNotMatch(pdfNumber(0.000001234, 8), /e/i, 'no exponent');
});

test('numbers that cannot be written are refused', () => {
  assert.equal(codeOf(() => pdfNumber(NaN)), 'invalid');
  assert.equal(codeOf(() => pdfNumber(Infinity)), 'invalid');
  assert.equal(codeOf(() => pdfNumber(-Infinity)), 'invalid');
  assert.equal(codeOf(() => pdfNumber(1e21)), 'invalid');
  assert.equal(codeOf(() => pdfNumber('12' as unknown as number)), 'invalid');
});

// ------------------------------------------------------------------- values

test('names escape everything that is no plain character', () => {
  assert.equal(serialize(name('Type')), '/Type');
  assert.equal(serialize(name('A B')), '/A#20B');
  assert.equal(serialize(name('a/b(c)#<>[]{}%')), '/a#2fb#28c#29#23#3c#3e#5b#5d#7b#7d#25');
  assert.equal(serialize(name('ş')), '/#c5#9f');
  assert.equal(serialize(name('😀')), '/#f0#9f#98#80');
  assert.equal(codeOf(() => serialize(name('a\0b'))), 'invalid');
});

test('literal strings escape parentheses, the backslash and every byte outside printable ASCII', () => {
  assert.equal(serialize(new PdfString(latin1('plain text'))), '(plain text)');
  assert.equal(serialize(new PdfString(latin1('a(b)c\\d'))), '(a\\(b\\)c\\\\d)');
  assert.equal(serialize(new PdfString(Uint8Array.from([0, 10, 13, 127, 128, 255]))), '(\\000\\012\\015\\177\\200\\377)');
  assert.equal(serialize(new PdfString(latin1(') endobj'))), '(\\) endobj)');
  assert.equal(serialize(new PdfString(Uint8Array.from([0xde, 0xad, 0x0f]), true)), '<dead0f>');
});

test('text strings are ASCII where they can be and UTF-16BE where they cannot', () => {
  assert.equal(serialize(pdfText('Report 2026')), '(Report 2026)');
  assert.deepEqual(Array.from(pdfText('Ağ').bytes), [0xfe, 0xff, 0, 0x41, 0x01, 0x1f]);
  assert.deepEqual(Array.from(pdfText('😀').bytes), [0xfe, 0xff, 0xd8, 0x3d, 0xde, 0x00], 'a pair of surrogates stays a pair');
  assert.deepEqual(Array.from(pdfText('a\ud83d').bytes), [0xfe, 0xff, 0, 0x61, 0xff, 0xfd], 'a lone surrogate becomes U+FFFD');
  assert.equal(serialize(pdfText('a\0b\u0007c\td\ne')), '(abc d e)', 'controls are left out, tabs and line ends are spaces');
});

test('dates are written in universal time', () => {
  const date = pdfDate(new Date(Date.UTC(2026, 9, 13, 8, 5, 9)));
  assert.ok(date, 'a date is written');
  assert.equal(serialize(date), '(D:20261013080509Z)');
  assert.equal(pdfDate(new Date(NaN)), undefined);
  assert.equal(pdfDate('2026' as unknown as Date), undefined);
});

test('dictionaries and arrays are written through, and undefined entries are left out', () => {
  assert.equal(serialize({ Type: name('Page'), Box: [0, 0, 595.28, 841.89], Ok: true, No: false, None: null, Skip: undefined, To: new PdfRef(7) }), '<</Type /Page /Box [0 0 595.28 841.89] /Ok true /No false /None null /To 7 0 R>>');
  assert.equal(serialize({ 'a b': 1 }), '<</a#20b 1>>');
  assert.equal(codeOf(() => serialize(new PdfRef(0))), 'invalid');
  assert.equal(codeOf(() => serialize(new PdfRef(1.5))), 'invalid');
  assert.equal(codeOf(() => serialize('text' as unknown as PdfName)), 'invalid');
  let deep: unknown[] = [];
  for (let i = 0; i < 100; i++) deep = [deep];
  assert.equal(codeOf(() => serialize(deep as never)), 'invalid');
});

// ------------------------------------------------------------------- writer

test('the cross reference table points at every object', () => {
  const file = minimal(new PdfWriter());
  const source = text(file);
  assert.ok(source.startsWith('%PDF-1.4\n%'), 'the file starts with the header');
  assert.ok(file[10] > 127 && file[11] > 127, 'the second line holds bytes above 127');
  assert.ok(source.endsWith('%%EOF\n'), 'the file ends with %%EOF');
  const start = Number((/startxref\n(\d+)\n%%EOF\n$/.exec(source) as RegExpExecArray)[1]);
  assert.equal(source.substr(start, 9), 'xref\n0 4\n');
  const entries = source.substr(start + 9, 80);
  assert.equal(entries.substr(0, 20), '0000000000 65535 f \n');
  for (let id = 1; id <= 3; id++) {
    const entry = entries.substr(id * 20, 20);
    assert.match(entry, /^\d{10} 00000 n \n$/, `entry ${id} is 20 bytes long`);
    assert.equal(source.substr(Number(entry.slice(0, 10)), `${id} 0 obj\n`.length), `${id} 0 obj\n`, `entry ${id} points at its object`);
  }
  const checked = checkPdf(file);
  assert.equal(checked.pages.length, 1);
  assert.deepEqual([checked.pages[0].width, checked.pages[0].height], [200, 100]);
});

test('the length of a stream is the length of its bytes, also when they look like the end of a stream', () => {
  const writer = new PdfWriter();
  const pages = writer.reserve();
  const data = latin1('q\nendstream\nendobj\n1 0 obj\n\0\xff');
  const content = writer.addStream({}, data);
  const page = writer.add({ Type: name('Page'), Parent: pages, MediaBox: [0, 0, 10, 10], Resources: {}, Annots: [content] });
  writer.add({ Type: name('Pages'), Kids: [page], Count: 1 }, pages);
  writer.finish({ root: writer.add({ Type: name('Catalog'), Pages: pages }) });
  const file = joinParts(writer.takeParts());
  const checked = checkPdf(file);
  const stream = checked.objects.get(content.id);
  assert.ok(stream && stream.stream, 'the stream is read back');
  assert.deepEqual(Array.from(stream.stream), Array.from(data));
  assert.equal((stream.value as { Length: number }).Length, data.length);
});

test('an object can be reserved and written later, but not twice and not never', () => {
  const writer = new PdfWriter();
  const later = writer.reserve();
  const first = writer.add({ Next: later });
  assert.equal(first.id, 2);
  assert.equal(codeOf(() => writer.finish({ root: first })), 'invalid', 'a reserved object that was never written');
  writer.add(12, later);
  assert.equal(codeOf(() => writer.add(13, later)), 'invalid', 'the same object twice');
  assert.equal(codeOf(() => writer.add(13, new PdfRef(99))), 'invalid', 'a number nobody reserved');
  writer.finish({ root: first });
  assert.equal(codeOf(() => writer.add(1)), 'invalid', 'nothing after the end');
});

test('the parts of the file are handed out as they are written', () => {
  const writer = new PdfWriter();
  const head = writer.takeParts();
  assert.equal(text(joinParts(head)).slice(0, 8), '%PDF-1.4');
  assert.deepEqual(writer.takeParts(), [], 'nothing is handed out twice');
  writer.addStream({}, new Uint8Array(1000));
  const parts = writer.takeParts();
  assert.equal(joinParts(parts).length, writer.size - joinParts(head).length);
  assert.ok(parts.some((part) => part.length === 1000), 'the bytes of a stream are a part of their own: they are not copied into a larger buffer');
});

test('a file beyond the limit is refused, and the limit is below 4 GB', () => {
  assert.equal(MAX_PDF_BYTES, 0xffffffff);
  const writer = new PdfWriter({ maxBytes: 300 });
  writer.add({ A: 1 });
  assert.equal(codeOf(() => writer.addStream({}, new Uint8Array(400))), 'too-large');
  // What was refused left nothing behind: the file can still be finished.
  const size = writer.size;
  assert.ok(size < 300, 'nothing of the refused stream was written');
  assert.equal(codeOf(() => new PdfWriter({ maxBytes: 5 })), 'too-large');
  const beyond = new PdfWriter({ maxBytes: MAX_PDF_BYTES * 4 });
  assert.equal(codeOf(() => beyond.addStream({}, { length: MAX_PDF_BYTES } as Uint8Array)), 'too-large', 'a limit above 4 GB is 4 GB');
});

test('offsets of ten digits are written as they are', () => {
  // The writer is told that much was written already: no test can hold 2 GB.
  const writer = new PdfWriter();
  (writer as unknown as { offset: number }).offset = 3999999990;
  const ref = writer.add({ A: 1 });
  (writer as unknown as { offset: number }).offset = 4100000000;
  writer.finish({ root: ref });
  const tail = text(joinParts(writer.takeParts()));
  assert.match(tail, /\n3999999990 00000 n \n/);
  assert.match(tail, /startxref\n4100000000\n%%EOF\n$/);
});

// -------------------------------------------------------------- compression

test('zlib data of the engine itself is read by another inflater, as the built-in one is', async () => {
  const data = new Uint8Array(70000);
  for (let i = 0; i < data.length; i++) data[i] = (i * 7 + (i >> 5)) & 255;
  assert.equal(adler32(latin1('Wikipedia')), 0x11e60398);
  assert.equal(adler32(new Uint8Array(0)), 1);
  const own = zlibOwn(data);
  assert.deepEqual(Array.from(inflateSync(own)), Array.from(data));
  assert.deepEqual(Array.from(inflateSync(zlibOwn(new Uint8Array(0)))), []);
  assert.deepEqual(Array.from(inflateSync(await zlibDeflate(data))), Array.from(data));
  assert.deepEqual(Array.from(await zlibDeflate(data, false)), Array.from(own), 'without the built-in stream the own code is used');
});

// --------------------------------------------------------------------- font

test('the glyphless font is a sound TrueType file of two empty glyphs', () => {
  const font = glyphlessFont();
  assert.ok(font.length < 1024, `the font is small (${font.length} bytes)`);
  const checked = checkTrueType(font);
  assert.equal(checked.glyphs, 2);
  assert.equal(checked.unitsPerEm, 1000);
  assert.equal(checked.advance, GLYPH_WIDTH);
  for (const tag of ['OS/2', 'cmap', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'name', 'post']) assert.ok(checked.tables.indexOf(tag) >= 0, `the font has a ${tag} table`);
  // Every character is mapped to glyph 1: a "many to one" table over all of Unicode.
  const cmap = checked.table('cmap');
  const view = new DataView(cmap.buffer, cmap.byteOffset, cmap.byteLength);
  assert.equal(view.getUint16(2), 1, 'one encoding record');
  const sub = view.getUint32(8);
  assert.equal(view.getUint16(sub), 13, 'format 13');
  assert.deepEqual([view.getUint32(sub + 12), view.getUint32(sub + 16), view.getUint32(sub + 20), view.getUint32(sub + 24)], [1, 0, 0x10ffff, 1]);
  assert.deepEqual(Array.from(glyphlessFont()), Array.from(font), 'the same bytes every time');
});

test('the glyphless font is taken by FreeType', { skip: !canRun('fc-query') && 'fc-query is not installed' }, () => {
  const folder = join(process.cwd(), `tmp-pdf-font-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  try {
    const path = join(folder, 'glyphless.ttf');
    writeFileSync(path, glyphlessFont());
    const out = execFileSync('fc-query', [path], { encoding: 'utf8' });
    assert.match(out, new RegExp(`family: "${FONT_NAME}"`), out);
    assert.match(out, /fontformat: "TrueType"/, out);
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
});

function canRun(tool: string): boolean {
  try {
    execFileSync(tool, ['--version'], { stdio: 'pipe' });
    return true;
  } catch {
    return false;
  }
}

test('every character gets a code of its own, in the order of its first use', () => {
  const font = new PdfTextFont();
  assert.deepEqual(font.encode('aba'), { hex: '000100020001', count: 3 });
  assert.deepEqual(font.encode('bç'), { hex: '00020003', count: 2 });
  assert.equal(font.codes, 3);
  assert.deepEqual(font.encode(''), { hex: '', count: 0 });
});

test('a character outside the BMP is one code, mapped to its pair of surrogates', () => {
  const font = new PdfTextFont();
  assert.deepEqual(font.encode('a😀b'), { hex: '000100020003', count: 3 });
  const map = parseToUnicode(font.toUnicode());
  assert.equal(map.get(1), 'a');
  assert.equal(map.get(2), '😀');
  assert.equal(map.get(3), 'b');
  assert.match(font.toUnicode(), /<0002> <d83dde00>/i);
});

test('halves of a surrogate pair, controls and noncharacters never get into the map', () => {
  const font = new PdfTextFont();
  // A lone high and a lone low surrogate, a NUL, a bell, a tab, a line feed, U+FFFE, a delete.
  const encoded = font.encode('a\ud83db\ude00c\0d\u0007e\tf\ng￾h\u007fi');
  const map = parseToUnicode(font.toUnicode());
  let out = '';
  for (let i = 0; i < encoded.hex.length; i += 4) out += map.get(parseInt(encoded.hex.substr(i, 4), 16));
  assert.equal(out, 'a�b�cde f g�hi');
  assert.equal(encoded.count, out.length);
  map.forEach((mapped, code) => {
    assert.ok(code >= 1 && code <= 0xffff, 'a code of two bytes');
    assert.doesNotMatch(mapped, /[\0-\u001f\u007f-\u009f]/, 'no control is mapped');
  });
});

test('Turkish letters and combining marks keep their characters', () => {
  const font = new PdfTextFont();
  const source = 'İstanbul ığüşöç ĞÜŞÖÇ é å';
  const encoded = font.encode(source);
  const map = parseToUnicode(font.toUnicode());
  let out = '';
  for (let i = 0; i < encoded.hex.length; i += 4) out += map.get(parseInt(encoded.hex.substr(i, 4), 16));
  assert.equal(out, source);
  assert.equal(encoded.count, source.length, 'a combining mark is a code of its own');
});

test('the ToUnicode map is written in blocks of at most 100 entries', () => {
  const font = new PdfTextFont();
  let source = '';
  for (let i = 0; i < 250; i++) source += String.fromCharCode(0x4e00 + i);
  font.encode(source);
  const cmap = font.toUnicode();
  assert.deepEqual(cmap.match(/\d+ beginbfchar/g), ['100 beginbfchar', '100 beginbfchar', '50 beginbfchar']);
  assert.equal(parseToUnicode(cmap).size, 250);
  assert.match(cmap, /\/CMapName \/Adobe-Identity-UCS def/);
});

test('beyond the codes of the font a character is U+FFFD rather than another one', () => {
  const font = new PdfTextFont();
  const limit = MAX_FONT_CODES;
  // Fill the font: every code point from U+10000 on is a character of its own.
  let filled = 0;
  for (let block = 0; filled < limit; block++) {
    let source = '';
    for (let i = 0; i < 4096 && filled < limit; i++, filled++) source += String.fromCodePoint(0x10000 + block * 4096 + i);
    font.encode(source);
  }
  assert.equal(font.codes, limit);
  const more = font.encode('一丁');
  assert.equal(font.codes, limit, 'no code is added');
  const map = parseToUnicode(font.toUnicode());
  assert.equal(map.get(parseInt(more.hex.slice(0, 4), 16)), '�');
  assert.equal(more.hex.slice(0, 4), more.hex.slice(4, 8));
});

// ------------------------------------------------------------------- images

/** The smallest JPEG header there is to read: SOI, a quantisation table that is skipped, the frame header. */
function jpegOf(width: number, height: number, components: number, marker = 0xc0): Uint8Array {
  const frame = [0xff, marker, 0, 8 + components * 3, 8, height >> 8, height & 255, width >> 8, width & 255, components];
  for (let i = 0; i < components; i++) frame.push(i + 1, 0x11, 0);
  return Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 4, 1, 2, 0xff, 0xff, 0xdb, 0, 3, 9, ...frame, 0xff, 0xd9]);
}

test('the size and the components of a JPEG are read from its frame header', () => {
  assert.deepEqual(jpegInfo(jpegOf(1240, 1754, 3)), { width: 1240, height: 1754, components: 3, adobe: false });
  assert.deepEqual(jpegInfo(jpegOf(8, 9, 1, 0xc2)), { width: 8, height: 9, components: 1, adobe: false });
  assert.deepEqual(jpegInfo(jpegOf(8, 9, 4)), { width: 8, height: 9, components: 4, adobe: false });
});

test('what is no JPEG is refused without reading beyond its end', () => {
  assert.equal(codeOf(() => jpegInfo(new Uint8Array(0))), 'invalid');
  assert.equal(codeOf(() => jpegInfo(latin1('%PDF-1.4'))), 'invalid');
  assert.equal(codeOf(() => jpegInfo(Uint8Array.from([0xff, 0xd8]))), 'invalid');
  assert.equal(codeOf(() => jpegInfo(Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0xff, 0xff]))), 'invalid', 'a segment longer than the file');
  assert.equal(codeOf(() => jpegInfo(Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 0, 0xff, 0xe0, 0, 0]))), 'invalid', 'a segment of no length');
  assert.equal(codeOf(() => jpegInfo(jpegOf(0, 9, 3))), 'invalid', 'no width');
  assert.equal(codeOf(() => jpegInfo(jpegOf(8, 9, 2))), 'invalid', 'two components');
  assert.equal(codeOf(() => jpegInfo(jpegOf(8, 9, 3).subarray(0, 22))), 'invalid', 'cut in the frame header');
  // Markers without end: the search gives up.
  const endless = new Uint8Array(70000).fill(0xff);
  endless[1] = 0xd8;
  assert.equal(codeOf(() => jpegInfo(endless)), 'invalid');
});

test('pixels become grey samples where they are grey, a fill where they are one colour', () => {
  const rgba = (pixels: number[][]): Uint8Array => Uint8Array.from(pixels.reduce((all, pixel) => all.concat(pixel, 255), [] as number[]));
  assert.deepEqual(imageOfPixels(rgba([[255, 255, 255], [255, 255, 255]]), 2, 1), { fill: [255, 255, 255] });
  assert.deepEqual(imageOfPixels(rgba([[1, 2, 3], [1, 2, 3], [1, 2, 3], [1, 2, 3]]), 2, 2), { fill: [1, 2, 3] });
  const grey = imageOfPixels(rgba([[0, 0, 0], [200, 200, 200]]), 2, 1);
  assert.deepEqual(grey, { image: { kind: 'samples', data: Uint8Array.from([0, 200]), width: 2, height: 1, colors: 1 } });
  const colour = imageOfPixels(rgba([[0, 0, 0], [200, 201, 200]]), 1, 2);
  assert.deepEqual(colour, { image: { kind: 'samples', data: Uint8Array.from([0, 0, 0, 200, 201, 200]), width: 1, height: 2, colors: 3 } });
  assert.equal(codeOf(() => imageOfPixels(new Uint8Array(7), 2, 1)), 'invalid', 'fewer pixels than the size says');
  assert.equal(codeOf(() => imageOfPixels(new Uint8Array(0), 0, 0)), 'invalid');
});

// ---------------------------------------------------------------- documents

const A4: [number, number] = [595.28, 841.89];

/** A page with a small picture and the given lines of text. */
function pageOf(lines: string[], options: { size?: [number, number]; image?: PdfPage['image'] } = {}): PdfPage {
  const size = options.size || A4;
  // The text is measured in px of a page 96 to the inch.
  const width = size[0] / 0.75;
  const samples = new Uint8Array(4 * 6 * 3);
  for (let i = 0; i < samples.length; i++) samples[i] = (i * 37) & 255;
  return {
    width: size[0],
    height: size[1],
    image: options.image || { kind: 'samples', data: samples, width: 4, height: 6, colors: 3 },
    text: { width, runs: lines.map((line, index) => ({ text: line, x: 96, y: 120 + index * 24, width: line.length * 8, size: 16 })) },
  };
}

test('a document of pages with pictures and text is a sound PDF', async () => {
  const file = await buildPdf([pageOf(['First line', 'Second line']), pageOf(['Another page'], { size: [612, 792] })], { title: 'Ölçüm raporu', author: 'A (B) \\ C', date: new Date(Date.UTC(2026, 9, 13, 12, 0, 0)) });
  const checked = checkPdf(file);
  assert.equal(checked.version, '1.4');
  assert.equal(checked.pages.length, 2);
  assert.deepEqual([checked.pages[0].width, checked.pages[0].height], A4);
  assert.deepEqual([checked.pages[1].width, checked.pages[1].height], [612, 792]);
  assert.equal(checked.info.Title, 'Ölçüm raporu');
  assert.equal(checked.info.Author, 'A (B) \\ C');
  assert.equal(checked.info.Creator, 'BYM Office');
  assert.equal(checked.info.Producer, 'BYM Office');
  assert.equal(checked.info.CreationDate, 'D:20261013120000Z');
  // The picture fills the page.
  const image = checked.pages[0].images[0];
  assert.deepEqual([image.width, image.height, image.colorSpace, image.filter], [4, 6, 'DeviceRGB', 'FlateDecode']);
  assert.deepEqual(image.matrix, [595.28, 0, 0, 841.89, 0, 0]);
  // The text is there, invisible, where the runs are: 96 px from the left is 72 pt, 120 px from the top.
  const first = checked.pages[0].texts[0];
  assert.equal(first.text, 'First line');
  assert.equal(first.mode, 3, 'the text is neither filled nor stroked');
  assert.equal(first.x, 72);
  assert.ok(Math.abs(first.y - (841.89 - 90)) < 0.01, `the baseline is 90 pt below the top (${first.y})`);
  assert.equal(first.size, 12);
  assert.ok(Math.abs(first.advance - 60) < 0.05, `ten characters take the 80 px (60 pt) of the run (${first.advance})`);
  assert.match(checked.pages[0].text, /First line\s*Second line/);
  assert.equal(checked.pages[1].text.trim(), 'Another page');
  assert.equal(checked.fonts.length, 1, 'one font for all pages');
  assert.equal(checked.fonts[0].baseFont, FONT_NAME);
});

test('the same document gives the same bytes, and another one another identifier', async () => {
  const options = { title: 'T', date: new Date(Date.UTC(2026, 0, 2)) };
  const one = await buildPdf([pageOf(['a'])], options);
  const two = await buildPdf([pageOf(['a'])], options);
  assert.deepEqual(Array.from(one), Array.from(two));
  const other = await buildPdf([pageOf(['b'])], options);
  const idOf = (file: Uint8Array) => (/\/ID \[<([0-9a-f]{32})> <\1>\]/.exec(text(file)) as RegExpExecArray)[1];
  assert.notEqual(idOf(one), idOf(other));
});

test('a document is written page by page, and its parts can be taken while it grows', async () => {
  const document = new PdfDocument({ date: new Date(0) });
  const parts: Uint8Array[] = [];
  for (let i = 0; i < 3; i++) {
    await document.addPage(pageOf([`Page ${i + 1}`]));
    const taken = document.takeParts();
    assert.ok(taken.length > 0, `page ${i + 1} was written when it was added`);
    parts.push(...taken);
  }
  assert.equal(document.pageCount, 3);
  await document.finish();
  parts.push(...document.takeParts());
  const file = joinParts(parts);
  assert.equal(file.length, document.size);
  assert.deepEqual(checkPdf(file).pages.map((page) => page.text.trim()), ['Page 1', 'Page 2', 'Page 3']);
  assert.equal(await rejection(document.addPage(pageOf(['late']))), 'invalid', 'no page after the end');
});

test('a document without a page is refused', async () => {
  assert.equal(await rejection(buildPdf([])), 'empty');
  assert.equal(await rejection(new PdfDocument().finish()), 'empty');
});

test('a page of one colour holds no picture, and a page without text no font', async () => {
  const file = await buildPdf([{ width: 300, height: 200, fill: [255, 255, 255] }, { width: 300, height: 200, fill: [255, 0, 128] }, { width: 300, height: 200 }], { date: new Date(0) });
  const checked = checkPdf(file);
  assert.equal(checked.pages.length, 3);
  assert.deepEqual(checked.pages.map((page) => page.images.length), [0, 0, 0]);
  assert.equal(checked.pages[0].content, '', 'white paper is not painted at all');
  assert.match(checked.pages[1].content, /^1 0 0\.502 rg 0 0 300 200 re f$/);
  assert.equal(checked.fonts.length, 0);
  assert.doesNotMatch(text(file), /FontFile2/, 'the font is left out of a file without text');
  assert.ok(file.length < 1200, `three empty pages are small (${file.length} bytes)`);
});

test('grey samples, a soft mask and a JPEG are written as what they are', async () => {
  const jpeg = jpegOf(16, 8, 3);
  const file = await buildPdf(
    [
      { width: 100, height: 100, image: { kind: 'samples', data: Uint8Array.from([0, 64, 128, 255]), width: 2, height: 2, colors: 1, alpha: Uint8Array.from([255, 0, 255, 0]) } },
      { width: 100, height: 100, image: { kind: 'jpeg', data: jpeg, width: 16, height: 8, components: 3 } },
      { width: 100, height: 100, image: { kind: 'jpeg', data: jpegOf(16, 8, 1), width: 16, height: 8, components: 1 } },
    ],
    { date: new Date(0) },
  );
  const checked = checkPdf(file);
  const grey = checked.pages[0].images[0];
  assert.deepEqual([grey.colorSpace, grey.filter, Array.from(grey.samples as Uint8Array)], ['DeviceGray', 'FlateDecode', [0, 64, 128, 255]]);
  assert.ok(grey.mask, 'the image has a soft mask');
  assert.deepEqual(Array.from(grey.mask.samples as Uint8Array), [255, 0, 255, 0]);
  const photo = checked.pages[1].images[0];
  assert.deepEqual([photo.colorSpace, photo.filter, photo.width, photo.height], ['DeviceRGB', 'DCTDecode', 16, 8]);
  assert.deepEqual(Array.from(photo.data), Array.from(jpeg), 'the bytes of the JPEG are in the file as they were');
  assert.equal(checked.pages[2].images[0].colorSpace, 'DeviceGray');
});

test('a picture that does not go with its size is refused', async () => {
  assert.equal(await rejection(buildPdf([{ width: 10, height: 10, image: { kind: 'samples', data: new Uint8Array(5), width: 2, height: 2, colors: 3 } }])), 'invalid');
  assert.equal(await rejection(buildPdf([{ width: 10, height: 10, image: { kind: 'samples', data: new Uint8Array(4), width: 2, height: 2, colors: 1, alpha: new Uint8Array(3) } }])), 'invalid');
  assert.equal(await rejection(buildPdf([{ width: 10, height: 10, image: { kind: 'jpeg', data: jpegOf(16, 8, 3), width: 15, height: 8, components: 3 } }])), 'invalid', 'the JPEG says another size');
  assert.equal(await rejection(buildPdf([{ width: 0, height: 10 }])), 'invalid');
  assert.equal(await rejection(buildPdf([{ width: NaN, height: 10 }])), 'invalid');
  assert.equal(await rejection(buildPdf([{ width: 20000, height: 10 }])), 'invalid', 'a page beyond 200 inches');
});

test('text of a document cannot become syntax of the file', async () => {
  const hostile = [') Tj ET', '\\', '%comment', 'endstream', 'endobj', 'a\0b', '\ud83d', '<</Type /Catalog>>', '(((', '\r\n\r\nstream'];
  const file = await buildPdf([pageOf(hostile)], { title: ') >> endobj 99 0 obj << /Type /Catalog', author: '\\) \0 %', date: new Date(0) });
  const checked = checkPdf(file);
  assert.deepEqual(
    checked.pages[0].texts.map((run) => run.text).filter((run) => run !== ' '),
    [') Tj ET', '\\', '%comment', 'endstream', 'endobj', 'ab', '�', '<</Type /Catalog>>', '(((', '    stream'],
  );
  assert.equal(checked.info.Title, ') >> endobj 99 0 obj << /Type /Catalog');
  assert.equal(checked.info.Author, '\\)  %');
  // In the content stream the text is hexadecimal digits and nothing else.
  assert.doesNotMatch(checked.pages[0].content, /endstream|comment|Catalog/);
  for (const run of checked.pages[0].content.match(/<[^>]*> Tj/g) || []) assert.match(run, /^<([0-9a-f]{4})+> Tj$/);
});

test('the file is the same with the compression of the engine as with the built-in one', async () => {
  const pages = [pageOf(['Compressed twice'])];
  const native = checkPdf(await buildPdf(pages, { date: new Date(0) }));
  const own = checkPdf(await buildPdf(pages, { date: new Date(0), nativeCompression: false }));
  assert.equal(own.pages[0].content, native.pages[0].content);
  assert.deepEqual(Array.from(own.pages[0].images[0].samples as Uint8Array), Array.from(native.pages[0].images[0].samples as Uint8Array));
});

test('a document beyond the limit on its size is refused', async () => {
  const big: PdfPage = { width: 100, height: 100, image: { kind: 'jpeg', data: Uint8Array.from([...jpegOf(16, 8, 3).subarray(0, 33), ...new Uint8Array(5000), 0xff, 0xd9]), width: 16, height: 8, components: 3 } };
  assert.equal(await rejection(buildPdf([big, big, big], { maxBytes: 12000 })), 'too-large');
});

test('five hundred pages are written in little time and little memory', async () => {
  const started = Date.now();
  const document = new PdfDocument({ date: new Date(0), title: 'Five hundred pages' });
  // A grey picture of 200 x 280 per page, different on every page, and forty lines of text.
  const lines: string[] = [];
  for (let i = 0; i < 40; i++) lines.push(`Line ${i + 1}: the quick brown fox jumps over the lazy dog, ığüşöç İĞÜŞÖÇ`);
  const parts: Uint8Array[] = [];
  let largestPending = 0;
  for (let p = 0; p < 500; p++) {
    const samples = new Uint8Array(200 * 280);
    for (let i = 0; i < samples.length; i += 97) samples[i] = (p + i) & 255;
    await document.addPage(pageOf([`Page ${p + 1}`, ...lines], { image: { kind: 'samples', data: samples, width: 200, height: 280, colors: 1 } }));
    const taken = document.takeParts();
    let pending = 0;
    for (const part of taken) pending += part.length;
    largestPending = Math.max(largestPending, pending);
    parts.push(...taken);
  }
  await document.finish();
  parts.push(...document.takeParts());
  const file = joinParts(parts);
  const elapsed = Date.now() - started;
  // A guard against getting stuck, not a measurement: an idle machine takes about a second.
  assert.ok(elapsed < 60000, `500 pages took ${elapsed} ms`);
  assert.ok(largestPending < 200 * 280 * 2, `no page holds more than its own picture (${largestPending} bytes)`);
  const checked = checkPdf(file);
  assert.equal(checked.pages.length, 500);
  assert.equal(checked.pages[499].texts[0].text, 'Page 500');
  assert.equal(checked.fonts.length, 1);
});

// ------------------------------------------------------------------ Poppler

popplerTest('Poppler reads the document: pages, sizes, text, pictures and the font', async () => {
  const lines = ['İstanbul’da ığüşöç ĞÜŞÖÇ', 'Astral 😀 and 𝐀', 'é combining', ') Tj ET \\ %x endstream'];
  const jpeg = realJpeg();
  const file = await buildPdf(
    [pageOf(lines), pageOf(['Second page, US Letter'], { size: [612, 792] }), pageOf(['Third page'], { size: [841.89, 595.28], image: { kind: 'jpeg', data: jpeg, width: 8, height: 8, components: 3 } })],
    { title: 'Başlık 😀 (1)', author: 'Yazar', date: new Date(Date.UTC(2026, 9, 13, 12, 0, 0)) },
  );
  const report = poppler(file);
  assert.equal(report.complaints, '', 'no tool complains about the file');
  assert.equal(report.pages, 3);
  assert.deepEqual(report.sizes, [{ width: 595.28, height: 841.89 }, { width: 612, height: 792 }, { width: 841.89, height: 595.28 }]);
  assert.equal(report.info.Title, 'Başlık 😀 (1)');
  assert.equal(report.info.Author, 'Yazar');
  assert.equal(report.info.Creator, 'BYM Office');
  assert.equal(report.info['PDF version'], '1.4');
  assert.equal(report.info.Encrypted, 'no');
  assert.deepEqual(report.text[0].trim().split(/\s*\n/), lines, 'the text comes out line by line, with every character, a combining mark after its letter');
  assert.equal(report.text[1].trim(), 'Second page, US Letter');
  assert.equal(report.text[2].trim(), 'Third page');
  assert.deepEqual(report.images, [
    { page: 1, width: 4, height: 6, color: 'rgb', encoding: 'image' },
    { page: 2, width: 4, height: 6, color: 'rgb', encoding: 'image' },
    { page: 3, width: 8, height: 8, color: 'rgb', encoding: 'jpeg' },
  ]);
  assert.deepEqual(report.fonts, [{ name: FONT_NAME, type: 'CID TrueType', embedded: true, unicode: true }]);
});

/** A real JPEG of 8 x 8 grey pixels in three components, made by hand: one block, all coefficients zero. */
function realJpeg(): Uint8Array {
  const bytes: number[] = [0xff, 0xd8];
  // One quantisation table of ones.
  bytes.push(0xff, 0xdb, 0, 67, 0, ...new Array(64).fill(1));
  bytes.push(0xff, 0xc0, 0, 17, 8, 0, 8, 0, 8, 3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0);
  // Huffman tables with one code of one bit each: DC category 0, AC end of block.
  bytes.push(0xff, 0xc4, 0, 20, 0x00, 1, ...new Array(15).fill(0), 0);
  bytes.push(0xff, 0xc4, 0, 20, 0x10, 1, ...new Array(15).fill(0), 0);
  bytes.push(0xff, 0xda, 0, 12, 3, 1, 0x00, 2, 0x00, 3, 0x00, 0, 63, 0);
  // Three components x (DC 0, end of block) = six zero bits, padded with ones.
  bytes.push(0x03, 0xff, 0xd9);
  return Uint8Array.from(bytes);
}
