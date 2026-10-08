// Files that are cut short, damaged or made to hurt: the reader ends with
// what it has or with an error it names, and never hangs, overflows the
// stack or takes more memory than the file is worth.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import { readRtf } from '../src/formats/rtf/document';
import { paragraphText, type Table } from '../src/model/doc/types';
import { fakeMeasurer } from './doc-layout-fixtures';
import { bytesOf, doc, drawingsOf, fakePng, hexOf, itemsOf, paragraphs, paragraphsOf, read, tables, texts } from './rtf-fixtures';

// A document with a bit of everything, for cutting and damaging.
const SAMPLE =
  String.raw`{\rtf1\ansi\ansicpg1254\deff0\uc1{\fonttbl{\f0\froman\fcharset162{\*\panose 02020603050405020304}Times New Roman{\*\falt Times};}{\f1\fswiss Arial;}}` +
  String.raw`{\colortbl;\red255\green0\blue0;}{\stylesheet{\ql\f0\fs24 Normal;}{\s1\qc\b\f1\fs32\sbasedon0 heading 1;}{\*\cs10\additive Default Paragraph Font;}}` +
  '{\\*\\listtable{\\list\\listtemplateid1{\\listlevel\\levelnfc0\\leveljc0\\levelfollow0\\levelstartat1{\\leveltext\\' + "'02\\'00.;}{\\levelnumbers\\'01;}\\fi-360\\li720}\\listid7}}" +
  String.raw`{\*\listoverridetable{\listoverride\listid7\listoverridecount0\ls1}}{\info{\title T}}\paperw11906\paperh16838\margl1417\margr1417\widowctrl` +
  String.raw`\sectd{\header\pard\plain\qc Head {\field{\*\fldinst PAGE}{\fldrslt 1}}\par}{\footer\pard\plain Foot\par}` +
  String.raw`\pard\plain\s1\qc\b\f1\fs32 Title \u305?\'fd\par` +
  String.raw`{\listtext\pard\plain 1.\tab}\pard\plain\fi-360\li720\ls1 Item{\super\chftn}{\footnote\pard\plain{\super\chftn} Note}\par` +
  String.raw`\trowd\trgaph108\trleft-108\clbrdrt\brdrs\brdrw10 \cellx3000\clvmgf\cellx6000\pard\intbl A\cell B\cell\row` +
  String.raw`\pard\plain {\field{\*\fldinst HYPERLINK "https://example.org"}{\fldrslt link}} {\*\bkmkstart b}x{\*\bkmkend b}` +
  String.raw`{\pict\pngblip\picwgoal600\pichgoal300 ` +
  hexOf(fakePng(40, 20, 8)) +
  String.raw`}{\shp{\*\shpinst\shpleft0\shptop0\shpright2000\shpbottom1000{\sp{\sn shapeType}{\sv 202}}{\shptxt\pard\plain Box\par}}{\shprslt x}}\par\sect\sectd Last\par}`;

/** Reads bytes that may be anything: the reader either gives a model or refuses with its own error. */
function attempt(bytes: Uint8Array): void {
  try {
    const model = readRtf(bytes);
    assert.ok(Array.isArray(model.body.blocks));
  } catch (error) {
    assert.match(String((error as Error).message), /^Invalid rtf: /);
  }
}

test('the sample reads whole', () => {
  const model = read(SAMPLE);
  assert.deepEqual(texts(model), ['Title ' + String.fromCharCode(0x131, 0x131), 'Item', 'A', 'B', 'link x', '', 'Last']);
  assert.equal(model.warnings.length, 0);
  assert.equal(model.unknown.size, 0);
});

test('a file cut short anywhere is read as far as it goes', () => {
  const bytes = bytesOf(SAMPLE);
  for (let length = 5; length < bytes.length; length++) {
    const model = readRtf(bytes.subarray(0, length));
    assert.ok(Array.isArray(model.body.blocks));
  }
  // Shorter than its first word, it is not an RTF file.
  for (let length = 0; length < 5; length++) assert.throws(() => readRtf(bytes.subarray(0, length)), /Invalid rtf/);
});

test('what was read before the cut is kept: text, cells, fields, pictures, notes', () => {
  const cut = (source: string) => read(source);
  assert.deepEqual(texts(cut(String.raw`{\rtf1 one\par two {\b bo`)), ['one', 'two bo']);
  assert.deepEqual(texts(cut(String.raw`{\rtf1\trowd\cellx3000\cellx6000\pard\intbl A\cell B`)), ['A', 'B']);
  const field = cut(String.raw`{\rtf1 a{\field{\*\fldinst HYPERLINK "x"}{\fldrslt li`);
  assert.equal(paragraphText(paragraphs(field)[0]), 'ali');
  assert.deepEqual(itemsOf(paragraphs(field)[0]).filter((item) => item.type === 'fieldChar').map((item) => (item.type === 'fieldChar' ? item.charType : '')), ['begin', 'separate', 'end']);
  const picture = cut(String.raw`{\rtf1 a{\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(fakePng(4, 4, 0)).slice(0, 51));
  assert.equal(picture.media!.size, 1);
  assert.equal(Array.from(picture.media!.values())[0].bytes.length, 25);
  const note = cut(String.raw`{\rtf1 a\chftn{\footnote\pard\plain\chftn no`);
  assert.equal(paragraphText(paragraphsOf(note.footnotes!.items.get(1)!.blocks)[0]), 'no');
  const header = cut(String.raw`{\rtf1{\header\pard He`);
  assert.equal(paragraphText(paragraphsOf(header.headers.get('hdr:0:default')!.blocks)[0]), 'He');
  const box = cut(String.raw`{\rtf1 a{\shp{\*\shpinst\shpleft0\shpright100{\shptxt\pard In`);
  assert.equal(paragraphText(paragraphsOf(drawingsOf(paragraphs(box)[0])[0].textBoxes![0])[0]), 'In');
  // Cut inside a control word, a number, a byte.
  assert.deepEqual(texts(cut(String.raw`{\rtf1 a\pa`)), ['a']);
  assert.deepEqual(texts(cut(String.raw`{\rtf1 a\fs2`)), ['a']);
  assert.deepEqual(texts(cut('{\\rtf1 a\\' + "'e")), ['a']);
  assert.deepEqual(texts(cut('{\\rtf1 a\\')), ['a']);
});

test('damaged files: any byte changed, any stretch removed or doubled', () => {
  const bytes = bytesOf(SAMPLE);
  // A fixed sequence of numbers, so that a failure can be repeated.
  let seed = 12345;
  const random = (limit: number) => {
    seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
    return (seed >>> 8) % limit;
  };
  const interesting = [0x7b, 0x7d, 0x5c, 0x27, 0x2a, 0x00, 0xff, 0x30, 0x2d, 0x20, 0x3b];
  for (let round = 0; round < 1500; round++) {
    let copy = bytes.slice();
    const kind = random(4);
    if (kind === 0) {
      for (let k = 0; k < 1 + random(4); k++) copy[5 + random(copy.length - 5)] = interesting[random(interesting.length)];
    } else if (kind === 1) {
      for (let k = 0; k < 1 + random(4); k++) copy[5 + random(copy.length - 5)] = random(256);
    } else if (kind === 2) {
      const from = 5 + random(copy.length - 5);
      const length = random(200);
      const next = new Uint8Array(copy.length - Math.min(length, copy.length - from));
      next.set(copy.subarray(0, from));
      next.set(copy.subarray(Math.min(copy.length, from + length)), from);
      copy = next;
    } else {
      const from = 5 + random(copy.length - 5);
      const length = Math.min(random(300), copy.length - from);
      const next = new Uint8Array(copy.length + length);
      next.set(copy.subarray(0, from + length));
      next.set(copy.subarray(from), from + length);
      copy = next;
    }
    attempt(copy);
  }
});

test('a hundred thousand groups in each other do not overflow the stack', () => {
  const depth = 100000;
  const model = read('{\\rtf1 a' + '{'.repeat(depth) + 'b' + '}'.repeat(depth) + 'c\\par}');
  assert.deepEqual(texts(model), ['abc']);
  assert.match(model.warnings.join('\n'), /nested deeper/);
  // Never closed, and with formatting on the way in.
  const open = read('{\\rtf1 a' + '{\\b '.repeat(depth) + 'b');
  assert.deepEqual(texts(open), ['ab']);
  // More closing braces than opening ones: the document ends with its own.
  assert.deepEqual(texts(read('{\\rtf1 a}}}}}b')), ['a']);
});

test('formatting of groups beyond the depth the reader follows does not leak out of the last one it does', () => {
  const model = read('{\\rtf1 a' + '{'.repeat(1500) + '\\b b' + '}'.repeat(1500) + 'c\\par}');
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'abc');
  const last = p.children[p.children.length - 1];
  assert.ok(last.type === 'run' && last.props.bold === undefined);
});

test('destinations nested without end: fields, notes, shapes, tables, pictures', () => {
  const depth = 20000;
  const fields = read('{\\rtf1 a' + '{\\field{\\*\\fldinst x}{\\fldrslt '.repeat(depth) + 'deep' + '}}'.repeat(depth) + 'b\\par}');
  assert.equal(texts(fields).length, 1);
  const notes = read('{\\rtf1 a' + '\\chftn{\\footnote '.repeat(depth) + 'deep' + '}'.repeat(depth) + 'b\\par}');
  assert.ok(notes.footnotes!.items.size <= 64);
  assert.equal(paragraphText(paragraphs(notes)[0]), 'ab');
  const shapes = read('{\\rtf1 a' + '{\\shp{\\*\\shpinst{\\shptxt '.repeat(depth) + 'deep' + '}}}'.repeat(depth) + 'b\\par}');
  assert.equal(paragraphText(paragraphs(shapes)[0]), 'ab');
  const groups = read('{\\rtf1 a' + '{\\shpgrp{\\*\\shpinst'.repeat(depth) + '}}'.repeat(depth) + 'b\\par}');
  assert.equal(paragraphText(paragraphs(groups)[0]), 'ab');
  const pictures = read('{\\rtf1 a' + '{\\pict '.repeat(depth) + 'ff' + '}'.repeat(depth) + 'b\\par}');
  assert.equal(paragraphText(paragraphs(pictures)[0]), 'ab');
  const headers = read('{\\rtf1 a' + '{\\header '.repeat(depth) + 'h' + '}'.repeat(depth) + 'b\\par}');
  assert.equal(headers.headers.size, 1);
  const stylesheets = read('{\\rtf1 ' + '{\\stylesheet'.repeat(depth) + '{\\s1 x;}' + '}'.repeat(depth) + 'b\\par}');
  assert.deepEqual(texts(stylesheets), ['b']);
});

test('a table deeper than any real one is read at the depth the reader follows, and lays out', () => {
  const model = read(doc(String.raw`\trowd\cellx3000\pard\intbl\itap2000000000 deep\nestcell\nestrow\pard\intbl top\cell\row\pard x\par`));
  let depth = 0;
  let table: Table | undefined = tables(model)[0];
  // Each round goes one table further in.
  while (table && depth < 100) {
    depth++;
    const inner: Table | undefined = table.rows[0].cells[0].blocks.find((block): block is Table => block.type === 'table');
    table = inner;
  }
  assert.equal(depth, 32);
  assert.deepEqual(texts(model), ['deep', 'top', 'x']);
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
});

test('numbers beyond every range are cut, not trusted', () => {
  const big = '99999999999999999999';
  const model = read(
    doc(
      `\\paperw${big}\\margl-${big}\\deftab${big}\\fs${big}\\li${big}\\fi-${big}\\sb${big}\\sl-${big}\\tx${big}\\expndtw${big}\\charscalex${big}\\up${big}\\f${big}\\cf${big}\\s${big}\\ls${big}\\ilvl${big}\\cols${big}\\outlinelevel${big}` +
        `\\trowd\\trgaph${big}\\trleft-${big}\\trrh${big}\\cellx${big}\\cellx${big}\\pard\\intbl\\itap${big} a\\cell\\row\\pard\\u${big}?\\uc${big} b\\par`,
    ),
  );
  assert.ok(texts(model).length >= 2);
  assert.equal(model.body.section!.page!.width, 1000000);
  const json = JSON.stringify(model.body);
  assert.ok(json.indexOf('2147483647') < 0, 'a raw number of the file is in the model');
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
});

test('"\\bin" with a length the file does not have, and a negative one', () => {
  assert.deepEqual(texts(read(String.raw`{\rtf1 a\bin999999999 bcd}e`)), ['a']);
  assert.deepEqual(texts(read(String.raw`{\rtf1 a\bin-5 bcd\par}`)), ['abcd']);
  const picture = read(String.raw`{\rtf1 a{\pict\pngblip\picwgoal10\pichgoal10\bin4000000000 abc`);
  assert.equal(Array.from(picture.media!.values())[0].bytes.length, 3);
});

test('rows and paragraphs without end: bounded by the paragraphs a document may have', () => {
  const cells = '{\\rtf1 ' + '\\cell'.repeat(5000) + '}';
  assert.throws(() => read(cells, { maxParagraphs: 1000 }), /Invalid rtf: the document has more than 1000 paragraphs/);
  assert.equal(texts(read(cells)).length, 5000);
  const pars = '{\\rtf1 ' + '\\par'.repeat(3000) + '}';
  assert.throws(() => read(pars, { maxParagraphs: 2999 }), /more than 2999 paragraphs/);
  assert.equal(texts(read(pars, { maxParagraphs: 3000 })).length, 3000);
  // Paragraphs of headers and notes count too.
  assert.throws(() => read('{\\rtf1{\\header ' + '\\par'.repeat(50) + '}a}', { maxParagraphs: 20 }), /more than 20 paragraphs/);
});

test('lists of things without end are cut: cells of a row, tab stops, fonts, colours, styles, list levels', () => {
  const row = read(doc('\\trowd' + '\\cellx100'.repeat(5000) + '\\pard\\intbl a\\cell\\row\\pard x\\par'));
  assert.ok(tables(row)[0].grid.length <= 512);
  let stops = '';
  for (let i = 1; i <= 5000; i++) stops += '\\tx' + i * 5;
  assert.equal(paragraphs(read(doc(stops + ' a\\par')))[0].props.tabs!.length, 64);
  let fonts = '';
  for (let i = 0; i < 12000; i++) fonts += '{\\f' + i + ' F' + i + ';}';
  assert.equal(read('{\\rtf1{\\fonttbl' + fonts + '}a}').fonts.size, 10000);
  const colors = read('{\\rtf1{\\colortbl' + '\\red1;'.repeat(12000) + '}\\cf11999 a}');
  assert.deepEqual(paragraphs(colors)[0].children.length, 1);
  let styles = '';
  for (let i = 1; i <= 21000; i++) styles += '{\\s' + i + ' S' + i + ';}';
  assert.equal(read('{\\rtf1{\\stylesheet' + styles + '}a}').styles.styles.size, 20000);
  const levels = read('{\\rtf1{\\*\\listtable{\\list' + ('{\\listlevel\\levelnfc0{\\leveltext\\' + "'01x;}}").repeat(50) + '\\listid1}}{\\*\\listoverridetable{\\listoverride\\listid1\\ls1}}a}');
  assert.equal(levels.numbering.abstracts.get(0)!.levels.size, 9);
});

test('names from the file never reach the members every object has', () => {
  const body =
    String.raw`{\*\bkmkstart __proto__}a{\*\bkmkend __proto__}{\*\bkmkstart constructor}b{\*\bkmkend constructor}` +
    String.raw`{\shp{\*\shpinst{\sp{\sn __proto__}{\sv polluted}}{\sp{\sn constructor}{\sv x}}{\sp{\sn toString}{\sv y}}}}` +
    String.raw`\constructor\toString\valueOf\hasOwnProperty1 c{\*\constructor hidden}\par`;
  const sheet = String.raw`{\fonttbl{\f0\constructor\toString __proto__;}{\f1 constructor;}}{\stylesheet{\s1\constructor __proto__;}{\s2 hasOwnProperty;}}`;
  const model = read('{\\rtf1' + sheet + body + '}');
  assert.deepEqual(texts(model), ['abc']);
  assert.equal(({} as { polluted?: unknown }).polluted, undefined);
  assert.equal(Object.prototype.hasOwnProperty.call(Object.prototype, 'polluted'), false);
  assert.equal(typeof model.fonts.get('__proto__')!.family, 'undefined');
  assert.deepEqual(Array.from(model.fonts.keys()), ['__proto__', 'constructor']);
  assert.equal(model.unknown.get('rtf/constructor'), 3);
  assert.equal(model.unknown.get('rtf/toString'), 2);
  assert.equal(model.unknown.get('rtf/hasOwnProperty'), 1);
});

test('a long run of one character does not take quadratic time', () => {
  const started = Date.now();
  const source = '{\\rtf1 ' + '\\'.repeat(300000) + '{'.repeat(1) + "\\'".repeat(300000) + '\\u'.repeat(200000) + ';'.repeat(300000) + '}}';
  const model = read(source);
  assert.ok(Array.isArray(model.body.blocks));
  assert.ok(Date.now() - started < 50000, `took ${Date.now() - started} ms`);
});

test('a file of ten megabytes, half of it a picture, reads in a second or so', () => {
  const picture = fakePng(800, 600, 2500000 - 33);
  // Line ends every 128 digits, as writers put them.
  const hex = hexOf(picture);
  const lines: string[] = [];
  for (let i = 0; i < hex.length; i += 128) lines.push(hex.slice(i, i + 128));
  const paragraph = String.raw`\pard\plain\s1\ql\fi-360\li720\sa200\f1\fs22 {\rtlch\fcs1 \af1\afs22 \ltrch\fcs0 \insrsid1234 Lorem ipsum dolor sit amet, }{\rtlch\fcs1 \ab\af1 \ltrch\fcs0 \b\insrsid5678 consectetur}{\rtlch\fcs1 \af1 \ltrch\fcs0 \insrsid1234  adipiscing elit \u305\'fd \'fe\'f0.}\par` + '\r\n';
  const row = String.raw`\trowd\trgaph108\trleft-108\clbrdrt\brdrs\brdrw10 \cellx3000\clbrdrt\brdrs\brdrw10 \cellx6000\pard\intbl {\insrsid1 cell one}\cell {\insrsid1 cell two}\cell\row` + '\r\n';
  let body = '';
  const wanted = 10 * 1024 * 1024 - hex.length;
  const unit = paragraph.repeat(8) + row.repeat(2);
  body = unit.repeat(Math.ceil(wanted / unit.length));
  const source = doc(body + '{\\pict\\pngblip\\picwgoal6000\\pichgoal4500 ' + lines.join('\r\n') + '}\\par', String.raw`{\stylesheet{\ql\f0\fs24 Normal;}{\s1\ql\fi-360\li720\sa200\f1\fs22 \sbasedon0 List;}}`);
  const bytes = bytesOf(source);
  assert.ok(bytes.length >= 10 * 1024 * 1024);
  const started = Date.now();
  const model = readRtf(bytes);
  const elapsed = Date.now() - started;
  const media = Array.from(model.media!.values());
  assert.equal(media.length, 1);
  assert.equal(media[0].bytes.length, picture.length);
  assert.equal(media[0].bytes[picture.length - 1], picture[picture.length - 1]);
  const all = paragraphs(model);
  assert.ok(all.length > 20000);
  assert.equal(all[0].props.style, 'List');
  // Paragraphs that say the same share their properties.
  assert.equal(all[0].props, all[1].props);
  // A generous limit: a laptop takes about a quarter of a second.
  assert.ok(elapsed < 50000, `reading took ${elapsed} ms`);
});
