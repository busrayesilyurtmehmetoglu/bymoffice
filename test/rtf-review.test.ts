// Tests for defects found by the independent review of the RTF reader: each
// of them failed before its fix.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { resolveRunProps } from '../src/formats/docx/resolve';
import { DocLayout } from '../src/layout/doc/layout';
import { paragraphText, type DocModel, type FieldCharItem, type Paragraph, type Run, type Table } from '../src/model/doc/types';
import { fakeMeasurer, round, texts as pageTexts } from './doc-layout-fixtures';
import { doc, effectiveRun, itemsOf, paragraphs, paragraphsOf, read, runWith, runsOf, tables, texts } from './rtf-fixtures';

const chars = (...codes: number[]): string => String.fromCharCode(...codes);

// ---------------------------------------------------------------------------
// Fields

/** The run of each field mark of a paragraph, in order. */
const marks = (paragraph: Paragraph): Run[] => runsOf(paragraph).filter((run) => run.content.length === 1 && run.content[0].type === 'fieldChar');
const markKind = (run: Run): string => (run.content[0] as FieldCharItem).charType;

/** Font and size in effect for a run. */
function look(model: DocModel, paragraph: Paragraph, run: Run): [string | undefined, number | undefined] {
  const props = resolveRunProps(model, paragraph.props, run.props);
  return [props.fonts && props.fonts.ascii, props.size];
}

test('the marks of a field have the formatting of its code and of its result, not of the text around it', () => {
  // As Word writes a page number: the text around it is 12 pt Times, the field 8 pt Arial.
  const model = read(doc(String.raw`\pard\plain\f0\fs24 Page {\field{\*\fldinst {\f1\fs16 PAGE }}{\fldrslt {\f1\fs16 1}}} of\par`));
  const p = paragraphs(model)[0];
  assert.deepEqual(marks(p).map(markKind), ['begin', 'separate', 'end']);
  for (const mark of marks(p)) assert.deepEqual(look(model, p, mark), ['Arial', 16], markKind(mark));
  assert.deepEqual(look(model, p, runWith(p, 'Page')), ['Times New Roman', 24]);
  assert.deepEqual(look(model, p, runWith(p, ' of')), ['Times New Roman', 24]);

  // Code and result formatted apart: the mark that begins the field goes with the code, the other two with the result.
  const apart = read(doc(String.raw`\pard\plain\f0\fs24 {\field{\*\fldinst {\fs40 PAGE }}{\fldrslt {\f1\fs16 1}{\fs60 2}}}\par`));
  const q = paragraphs(apart)[0];
  assert.deepEqual(marks(q).map((mark) => look(apart, q, mark)), [
    ['Times New Roman', 40],
    ['Arial', 16],
    ['Arial', 16],
  ]);

  // No result: the mark that ends the field goes with the code. Neither: the text around the field.
  const bare = read(doc(String.raw`\pard\plain\f0\fs24 {\field{\*\fldinst {\fs40 PAGE }}}{\b\fs30{\field}}\par`));
  const r = paragraphs(bare)[0];
  assert.deepEqual(marks(r).map((mark) => look(bare, r, mark)), [
    ['Times New Roman', 40],
    ['Times New Roman', 40],
    ['Times New Roman', 30],
    ['Times New Roman', 30],
  ]);
});

test('field marks keep the formatting of the field when the paragraph turns out to have another style', () => {
  // The style is stated after the field: the runs are made again for it, the marks with what they were given.
  const sheet = String.raw`{\stylesheet{\s1\f0\fs40 Big;}}`;
  const model = read(doc(String.raw`\pard\plain\f0\fs24 {\field{\*\fldinst {\f1\fs16 PAGE }}{\fldrslt {\f1\fs16 1}}}\s1\par`, sheet));
  const p = paragraphs(model)[0];
  assert.equal(p.props.style, 'Big');
  for (const mark of marks(p)) assert.deepEqual(look(model, p, mark), ['Arial', 16]);
});

test('a page number field of an RTF file is laid out in the size and font of its result', () => {
  const model = read(doc(String.raw`\paperw6000\paperh6000\margl600\margr600\margt600\margb600\pard\plain\f0\fs24 Page {\field{\*\fldinst {\f1\fs16 PAGE }}{\fldrslt {\f1\fs16 9}}}\par`));
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  const number = pageTexts(layout.pages[0]).filter((item) => item.text === '1');
  assert.equal(number.length, 1);
  assert.deepEqual([number[0].font.family, round(number[0].font.size)], ['Arial', 10.67]);
});

// ---------------------------------------------------------------------------
// Fonts of symbols

test('in a font of symbols "\\uN" below 256 is a glyph of the font, as "\\\'hh" is', () => {
  // f2 is Symbol: 214 is its square root, whichever way the file writes it.
  assert.equal(texts(read(doc(String.raw`{\f2 \u214\'d6}\par`)))[0], chars(0xf0d6));
  assert.equal(texts(read(doc(String.raw`{\f2 \'d6\u214?}\par`)))[0], chars(0xf0d6, 0xf0d6));
  // What Word writes is in the private use area already.
  assert.equal(texts(read(doc('{\\f2 \\u-3913?\\u61623?}\\par')))[0], chars(0xf0b7, 0xf0b7));
  // In any other font the number is the character.
  assert.equal(texts(read(doc(String.raw`{\f1 \u214?\u305?}{\f0 \u65?}\par`)))[0], 'Öı' + 'A');
});

test('plain text in a font of symbols is glyphs of the font too; spaces and tabs stay what they are', () => {
  const model = read(doc(String.raw`a {\f2 ab c\tab d\'20e}b\par`));
  assert.equal(texts(model)[0], 'a ' + chars(0xf061, 0xf062) + ' ' + chars(0xf063) + '\t' + chars(0xf064) + ' ' + chars(0xf065) + 'b');
  assert.equal(effectiveRun(model, paragraphs(model)[0], chars(0xf061)).fonts!.ascii, 'Symbol');
});

test('the code of a field is not text of a font of symbols', () => {
  const model = read(doc(String.raw`{\f2{\field{\*\fldinst SYMBOL 183 \\f "Symbol"}{\fldrslt \'b7}}}\par`));
  const begin = itemsOf(paragraphs(model)[0]).find((item): item is FieldCharItem => item.type === 'fieldChar' && item.charType === 'begin')!;
  assert.equal(begin.field!.instruction, 'SYMBOL 183 \\f "Symbol"');
  assert.equal(begin.field!.result, chars(0xf0b7));
});

test('the bullet of a list level in a font of symbols is a glyph of the font however it is written', () => {
  const level = (text: string, font: string) => '{\\list\\listsimple{\\listlevel\\levelnfc23{\\leveltext' + text + ';}{\\levelnumbers;}' + font + ' }\\listid' + (font === '\\f2' ? 1 : 2) + '}';
  const lists = (text: string) =>
    '{\\*\\listtable' + level(text, '\\f2') + level(text, '\\f0') + '}{\\*\\listoverridetable{\\listoverride\\listid1\\ls1}{\\listoverride\\listid2\\ls2}}';
  const textsOf = (model: DocModel) => [0, 1].map((id) => model.numbering.abstracts.get(id)!.levels.get(0)!.text);
  // A number below 256, a letter and a byte: glyphs in Symbol, characters in Times.
  assert.deepEqual(textsOf(read(doc('a\\par', lists("\\'01\\u183 ?")))), [chars(0xf0b7), chars(0xb7)]);
  assert.deepEqual(textsOf(read(doc('a\\par', lists("\\'01o")))), [chars(0xf06f), 'o']);
  assert.deepEqual(textsOf(read(doc('a\\par', lists("\\'01\\'b7")))), [chars(0xf0b7), chars(0xb7)]);
});

// ---------------------------------------------------------------------------
// Associated fonts

const FONTS =
  String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\froman\fcharset0 Times New Roman;}{\f1\fswiss\fcharset162 Arial;}{\f2\fnil\fcharset2 Symbol;}` +
  String.raw`{\f13\fnil\fcharset134 SimSun;}{\f31505\fnil\fcharset0 Theme East;}{\f31506\fswiss\fcharset0 Calibri;}}`;

function fontsOf(body: string, text: string): (string | undefined)[] {
  const model = read(FONTS + body + '}');
  const fonts = effectiveRun(model, paragraphs(model)[0], text).fonts!;
  return [fonts.ascii, fonts.hAnsi, fonts.eastAsia];
}

test('a plain "\\fN" after associated fonts is the font of the text and replaces them', () => {
  assert.deepEqual(fontsOf(String.raw`\f31506\loch\af31506\hich\af31506\dbch\af31505 {\f1 plain}\par`, 'plain'), ['Arial', 'Arial', 'Arial']);
  // Without the group as well.
  assert.deepEqual(fontsOf(String.raw`\loch\af31506\hich\af31506\dbch\af31505\f1 plain\par`, 'plain'), ['Arial', 'Arial', 'Arial']);
  // The bytes are read as that font has them: a bullet of Symbol.
  const model = read(FONTS + String.raw`{\loch\af0\hich\af0\dbch\af0 \f2 \'b7}\par}`);
  assert.equal(texts(model)[0], chars(0xf0b7));
  assert.equal(effectiveRun(model, paragraphs(model)[0], chars(0xf0b7)).fonts!.ascii, 'Symbol');
});

test('what Word writes stays as it was: associated fonts first, then the font of the kind of text at hand', () => {
  // Latin text: the font of the high half and of double-byte text are associated, "\loch\fN" is the font of the text.
  assert.deepEqual(fontsOf(String.raw`\f1\hich\af1\dbch\af31505\loch\f1 text\par`, 'text'), ['Arial', 'Arial', 'Theme East']);
  assert.deepEqual(fontsOf(String.raw`\hich\af31506\dbch\af31505\loch\f31506 text\par`, 'text'), ['Calibri', 'Calibri', 'Theme East']);
  // Text of two bytes: "\dbch\fN" is its font, the Latin fonts are associated.
  assert.deepEqual(fontsOf(String.raw`\loch\af0\hich\af0\dbch\f13 x\par`, 'x'), ['Times New Roman', 'Times New Roman', 'SimSun']);
  // "\loch\fN" ends an associated font stated for the low half before.
  assert.deepEqual(fontsOf(String.raw`\loch\af31506\hich\af31506\loch\f1 text\par`, 'text'), ['Arial', 'Calibri', 'Arial']);
});

// ---------------------------------------------------------------------------
// What stands in for a "\u"

test('a control word that is a character counts among the characters that stand in for a "\\u"', () => {
  assert.equal(texts(read(doc(String.raw`\u305\emdash x\par`)))[0], 'ıx');
  assert.equal(texts(read(doc(String.raw`\uc2\u305\emdash\endash x\par`)))[0], 'ıx');
  assert.equal(texts(read(doc(String.raw`\uc3\u305\lquote\rdblquote\bullet x\par`)))[0], 'ıx');
  // Control symbols and bytes as well.
  assert.equal(texts(read(doc(String.raw`\uc4\u305\~\_\-\'e7x\par`)))[0], 'ıx');
  assert.equal(texts(read(doc(String.raw`\uc3\u305\{\\\}x\par`)))[0], 'ıx');
  // Without a "\u" before them they are what they are.
  assert.equal(texts(read(doc(String.raw`a\emdash b\~c\par`)))[0], 'a—b c');
});

test('what ends a paragraph, a cell or a line is never passed over after a "\\u"', () => {
  assert.deepEqual(texts(read(doc(String.raw`\u305\par x\par`))), ['ı', 'x']);
  assert.deepEqual(texts(read(doc(String.raw`\uc2\u305\par x\par`))), ['ı', 'x']);
  assert.deepEqual(texts(read(doc(String.raw`\u305\line x\u305\tab y\u305\page z\par`))), ['ı\nxı\tyı\nz']);
  assert.deepEqual(texts(read(doc(String.raw`\u305\sect x\par`))), ['ı', 'x']);
  // Text in front of a row mark is a cell of its own; the mark ends the row all the same.
  const table = tables(read(doc(String.raw`\trowd\cellx1000\cellx2000\intbl \u305\cell x\u305\cell\u305\row\intbl z\cell\row\pard y\par`)))[0];
  assert.deepEqual(table.rows.map((row) => row.cells.map((cell) => paragraphsOf(cell.blocks).map(paragraphText).join('|'))), [['ı', 'xı', 'ı'], ['z']]);
  // A control word that formats ends the count and does what it says.
  const model = read(doc(String.raw`\uc2\u305\b xy\par`));
  assert.equal(texts(model)[0], 'ıxy');
  assert.equal(runWith(paragraphs(model)[0], 'xy').props.bold, true);
});

// ---------------------------------------------------------------------------
// Tables in cells

test('the description of a nested row in front of its cells is the one of that row', () => {
  const body = String.raw`\trowd\cellx6000\pard\intbl{\*\nesttableprops\trowd\trrh300\cellx2000\nestrow}\itap2 in1\nestcell\pard\intbl\itap1 out\cell\row\pard after\par`;
  const outer = tables(read(doc(body)))[0];
  assert.deepEqual(outer.grid, [6000]);
  const inner = outer.rows[0].cells[0].blocks[0] as Table;
  assert.equal(inner.type, 'table');
  assert.deepEqual(inner.grid, [2000]);
  assert.equal(inner.rows[0].props.height, 300);
  assert.deepEqual(inner.rows[0].cells[0].props.width, { type: 'dxa', value: 2000 });

  // Every row of such a writer: each description is for the cells after it.
  const two =
    String.raw`\trowd\cellx6000\pard\intbl{\*\nesttableprops\trowd\cellx2000\nestrow}\itap2 a\nestcell{\*\nesttableprops\trowd\cellx1000\cellx3000\nestrow}b\nestcell c\nestcell` +
    String.raw`\pard\intbl\itap1 out\cell\row\pard after\par`;
  const second = tables(read(doc(two)))[0].rows[0].cells[0].blocks[0] as Table;
  assert.deepEqual(second.rows.map((row) => row.cells.map((cell) => (cell.props.width as { value: number }).value)), [[2000], [1000, 2000]]);
});

test('the description of a nested row after its cells, as Word writes it, is still the one of that row', () => {
  const body =
    String.raw`\trowd\cellx6000\pard\intbl\itap2 a\nestcell{\*\nesttableprops\trowd\cellx2000\nestrow}\pard\intbl\itap2 b\nestcell c\nestcell{\*\nesttableprops\trowd\cellx1000\cellx3000\nestrow}` +
    String.raw`\pard\intbl\itap1 out\cell\row\pard after\par`;
  const inner = tables(read(doc(body)))[0].rows[0].cells[0].blocks[0] as Table;
  assert.deepEqual(inner.rows.map((row) => row.cells.map((cell) => (cell.props.width as { value: number }).value)), [[2000], [1000, 2000]]);
});

// ---------------------------------------------------------------------------
// What a file may ask for in memory

/** Objects and arrays that can be reached from a value. */
function countObjects(root: unknown): number {
  const seen = new Set<object>();
  const stack: unknown[] = [root];
  while (stack.length) {
    const value = stack.pop();
    if (value === null || typeof value !== 'object' || seen.has(value as object)) continue;
    seen.add(value as object);
    if (value instanceof Map) {
      for (const [key, item] of value) stack.push(key, item);
    } else if (Array.isArray(value)) {
      for (const item of value) stack.push(item);
    } else if (!(value instanceof Uint8Array)) {
      for (const key of Object.keys(value as object)) stack.push((value as { [key: string]: unknown })[key]);
    }
  }
  return seen.size;
}

const countTables = (model: DocModel): number => {
  let count = 0;
  const walk = (blocks: DocModel['body']['blocks']) => {
    for (const block of blocks) {
      if (block.type !== 'table') continue;
      count++;
      for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(model.body.blocks);
  return count;
};

test('a file that jumps in and out of deep tables gets no more tables than its size pays for', () => {
  // 31 bytes that open 31 tables in tables, again and again.
  const unit = String.raw`\itap32 a\nestcell\itap1 b\cell `;
  const times = 4000;
  const source = doc(unit.repeat(times));
  const model = read(source);
  const allowed = Math.floor(source.length / 32) + 1024;
  assert.ok(countTables(model) <= allowed, `${countTables(model)} tables`);
  assert.ok(countObjects(model) <= source.length * 2, `${countObjects(model)} objects for ${source.length} bytes`);
  // No text is lost: what no table was opened for is read as paragraphs.
  const all = paragraphsOf(model.body.blocks).map(paragraphText);
  assert.equal(all.filter((text) => text === 'a').length, times);
  assert.equal(all.filter((text) => text === 'b').length, times);
  assert.ok(model.warnings.some((warning) => /more tables than the reader builds/.test(warning)));
});

test('"maxTables" says how many tables are built; a genuine file is far below what it is given', () => {
  const row = String.raw`\trowd\cellx1000\pard\intbl a\cell\row\pard b\par`;
  const three = doc(row.repeat(3));
  assert.equal(countTables(read(three)), 3);
  assert.equal(read(three).warnings.length, 0);
  const capped = read(three, { maxTables: 2 });
  assert.equal(countTables(capped), 2);
  assert.deepEqual(paragraphsOf(capped.body.blocks).map(paragraphText), ['a', 'b', 'a', 'b', 'a', 'b']);
  assert.deepEqual(capped.body.blocks.map((block) => block.type), ['table', 'paragraph', 'table', 'paragraph', 'paragraph', 'paragraph']);
  // Nested tables count each.
  const nested = doc(String.raw`\trowd\cellx3000\pard\intbl\itap2 in\nestcell{\*\nesttableprops\trowd\cellx2000\nestrow}\pard\intbl\itap1 out\cell\row\pard x\par`);
  assert.equal(countTables(read(nested)), 2);
  const flat = read(nested, { maxTables: 1 });
  assert.equal(countTables(flat), 1);
  assert.deepEqual(paragraphsOf(flat.body.blocks).map(paragraphText), ['in', 'out', 'x']);
  assert.deepEqual(paragraphsOf(read(nested, { maxTables: 0 }).body.blocks).map(paragraphText), ['in', 'out', 'x']);
  assert.equal(countTables(read(nested, { maxTables: 0 })), 0);
});

test('"maxBytes" refuses a file larger than the host wants to read', () => {
  const source = doc('a\\par');
  assert.equal(texts(read(source, { maxBytes: source.length }))[0], 'a');
  assert.throws(() => read(source, { maxBytes: source.length - 1 }), /^Error: Invalid rtf: the file is larger than \d+ bytes$/);
});

test('rows, cells and paragraphs without anything in them cost a few small objects each', () => {
  const times = 3000;
  const rows = read(doc(String.raw`\cell\row`.repeat(times)));
  const table = tables(rows)[0];
  assert.equal(table.rows.length, times);
  // Nothing is stated for them: they share their properties.
  assert.equal(table.rows[0].props, table.rows[1].props);
  assert.equal(table.rows[0].cells[0].props, table.rows[1].cells[0].props);
  assert.ok(countObjects(rows) <= times * 6 + 200, `${countObjects(rows)} objects for ${times} rows`);

  const empty = read(doc(String.raw`\par`.repeat(times)));
  assert.ok(countObjects(empty) <= times * 2 + 200, `${countObjects(empty)} objects for ${times} paragraphs`);
});

test('text whose formatting goes back and forth shares the properties of its runs', () => {
  const times = 3000;
  const model = read(doc(String.raw`\b a\b0 b`.repeat(times) + '\\par'));
  const runs = runsOf(paragraphs(model)[0]);
  assert.equal(runs.length, times * 2);
  assert.equal(runs[0].props, runs[2].props);
  assert.equal(runs[1].props, runs[3].props);
  // A run, its items and one text item.
  assert.ok(countObjects(model) <= times * 2 * 3 + 200, `${countObjects(model)} objects for ${times * 2} runs`);
  // Text in one formatting is one run, however many pieces it comes in.
  const pieces = read(doc(String.raw`{\b a}{\b b}\b c\b d\par`));
  assert.deepEqual(runsOf(paragraphs(pieces)[0]).map((run) => run.content.length), [1]);
  assert.equal(texts(pieces)[0], 'abcd');
});

test('sections described alike share one description', () => {
  const model = read(doc(String.raw`a\sect b\sect\sectd\pgwsxn5000 c\sect d\par`));
  const [a, b, c] = paragraphs(model);
  assert.ok(a.props.section);
  assert.equal(a.props.section, b.props.section);
  assert.equal(a.props, b.props);
  assert.notEqual(c.props.section, a.props.section);
  assert.equal(c.props.section!.page!.width, 5000);
  // What two sections that differ say alike, they share.
  assert.equal(c.props.section!.margins, a.props.section!.margins);
  const many = read(doc(String.raw`\sect`.repeat(2000)));
  assert.ok(countObjects(many) <= 2000 * 2 + 200, `${countObjects(many)} objects`);
});
