// Behaviour of the RTF reader that no test pinned down: a change to any of
// these lines went unnoticed by the tests there were.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { resolveParaProps } from '../src/formats/docx/resolve';
import { readRtf } from '../src/formats/rtf/document';
import { MAX_CELLS } from '../src/formats/rtf/tables';
import { RtfTokenizer, T_BIN, T_CLOSE, T_EOF, T_OPEN, T_WORD } from '../src/formats/rtf/tokenizer';
import { paragraphText, type DocModel, type FieldCharItem, type RangeMark } from '../src/model/doc/types';
import { bytesOf, doc, drawingsOf, effectivePara, effectiveRun, itemsOf, paragraphs, paragraphsOf, read, runWith, tables, texts } from './rtf-fixtures';

const chars = (...codes: number[]): string => String.fromCharCode(...codes);

// ---------------------------------------------------------------------------
// The tokenizer

test('the data of "\\bin" ends where the file ends, and so does the tokenizer', () => {
  const bytes = bytesOf(String.raw`{\pict\bin500 abc`);
  const tk = new RtfTokenizer(bytes);
  assert.equal(tk.next(), T_OPEN);
  assert.equal(tk.next(), T_WORD);
  assert.equal(tk.next(), T_BIN);
  // Not the 500 bytes the file asks for: the three it has.
  assert.deepEqual([tk.start, tk.stop, tk.pos], [bytes.length - 3, bytes.length, bytes.length]);
  assert.equal(tk.next(), T_EOF);
  assert.equal(tk.next(), T_EOF);

  // The same when the group is passed over.
  const skipped = new RtfTokenizer(bytesOf(String.raw`{\x\bin500 a}c`));
  skipped.next();
  skipped.skipGroup();
  assert.equal(skipped.pos, 14);
  assert.equal(skipped.next(), T_EOF);
  // Data that ends before the file: the brace after it closes the group.
  const whole = new RtfTokenizer(bytesOf(String.raw`{\x\bin2 }}}c`));
  whole.next();
  whole.skipGroup();
  assert.equal(whole.pos, 11);
  assert.equal(whole.next(), T_CLOSE);
});

// ---------------------------------------------------------------------------
// Characters

test('"\\uc" goes into the groups inside and comes back when its own group ends', () => {
  // Two characters stand in, in the group around and in the one inside.
  assert.equal(texts(read(doc(String.raw`\uc2{\u305 abc}\u305 abc\par`)))[0], 'ıcıc');
  // Three in the group, one again after it.
  assert.equal(texts(read(doc(String.raw`{\uc3\u305 abcd}\u305 abcd\par`)))[0], 'ıdıbcd');
  // Deeper still, and back out step by step.
  assert.equal(texts(read(doc(String.raw`{\uc2{\uc0\u305 a}\u305 abc}\u305 abc\par`)))[0], 'ıaıcıbc');
});

test('"\\plain" ends right-to-left text and the formatting of complex scripts', () => {
  const model = read(doc(String.raw`\rtlch a\plain b\par`));
  const p = paragraphs(model)[0];
  assert.equal(runWith(p, 'a').props.rtl, true);
  assert.equal(effectiveRun(model, p, 'b').rtl, undefined);

  // After "\fcs1" bold and size are those of complex scripts; "\plain" goes back to the others.
  const complex = read(doc(String.raw`\fcs1\b\fs40 a\plain\b\fs30 b\par`));
  const q = paragraphs(complex)[0];
  assert.deepEqual([effectiveRun(complex, q, 'a').bold, effectiveRun(complex, q, 'a').boldCs, effectiveRun(complex, q, 'a').size, effectiveRun(complex, q, 'a').sizeCs], [undefined, true, 24, 40]);
  assert.deepEqual([effectiveRun(complex, q, 'b').bold, effectiveRun(complex, q, 'b').boldCs, effectiveRun(complex, q, 'b').size], [true, undefined, 30]);

  // A kind of text named before it no longer waits for its font: "\afN" is the font of the text again.
  const kind = read(doc(String.raw`\hich\plain\af1 x\par`));
  assert.equal(effectiveRun(kind, paragraphs(kind)[0], 'x').fonts!.ascii, 'Arial');
});

test('"\\pard" ends the border and the tab stop being described', () => {
  // The width after "\pard" belongs to no border.
  const model = read(doc(String.raw`\brdrt\brdrs\pard\brdrw30 x\par`));
  assert.equal(effectivePara(model, paragraphs(model)[0]).borders, undefined);
  // Without it, it is the width of the border.
  const kept = read(doc(String.raw`\brdrt\brdrs\brdrw30 x\par`));
  assert.deepEqual(effectivePara(kept, paragraphs(kept)[0]).borders, { top: { style: 'single', size: 12 } });
  // The alignment and the leader said for a tab stop that never came do not go to the next paragraph's.
  const tabs = read(doc(String.raw`\tqr\tldot\pard\tx1000 x\par`));
  assert.deepEqual(effectivePara(tabs, paragraphs(tabs)[0]).tabs, [{ pos: 1000, align: 'left' }]);
  const same = read(doc(String.raw`\tqr\tldot\tx1000 x\par`));
  assert.deepEqual(effectivePara(same, paragraphs(same)[0]).tabs, [{ pos: 1000, align: 'right', leader: 'dot' }]);
});

// ---------------------------------------------------------------------------
// Fields and stories

test('a second result of a field adds to the first: the field has one separator', () => {
  const model = read(doc(String.raw`{\field{\*\fldinst PAGE}{\fldrslt 1}{\fldrslt 2}}\par`));
  const marks = itemsOf(paragraphs(model)[0]).filter((item): item is FieldCharItem => item.type === 'fieldChar');
  assert.deepEqual(marks.map((item) => item.charType), ['begin', 'separate', 'end']);
  assert.deepEqual(marks[0].field, { instruction: 'PAGE', result: '12' });
});

test('a header inside a header is not a header: it is passed over', () => {
  const model = read(doc(String.raw`{\header a{\header b}{\footer c}d\par}x\par`));
  assert.equal(model.headers.size, 1);
  assert.equal(model.footers.size, 0);
  const header = Array.from(model.headers.values())[0];
  assert.deepEqual(paragraphsOf(header.blocks).map(paragraphText), ['ad']);
  assert.deepEqual(texts(model), ['x']);
});

test('a name gathered from the file is cut at 4096 characters', () => {
  // Characters that come in one at a time: the name stops growing at the limit.
  const name = 'n'.repeat(4000) + "\\'41".repeat(300);
  const model = read(doc('{\\*\\bkmkstart ' + name + '}x{\\*\\bkmkend ' + name + '}\\par'));
  const marks = paragraphs(model)[0].children.filter((inline): inline is RangeMark => inline.type === 'bookmarkStart' || inline.type === 'bookmarkEnd');
  assert.deepEqual(marks.map((mark) => mark.type), ['bookmarkStart', 'bookmarkEnd']);
  assert.equal((marks[0] as { name?: string }).name!.length, 4096);
  assert.equal((marks[0] as { name?: string }).name, 'n'.repeat(4000) + 'A'.repeat(96));
});

// ---------------------------------------------------------------------------
// Tables

test('a row describes no more cells than the reader takes; the cells beyond are cells without a description', () => {
  assert.equal(MAX_CELLS, 512);
  let row = '\\trowd';
  for (let i = 1; i <= 600; i++) row += '\\cellx' + i * 10;
  row += '\\pard\\intbl';
  for (let i = 0; i < 600; i++) row += ' c\\cell';
  const table = tables(read(doc(row + '\\row\\pard x\\par')))[0];
  assert.equal(table.grid.length, 512);
  assert.equal(table.rows[0].cells.length, 600);
  assert.deepEqual(table.rows[0].cells[511].props.width, { type: 'dxa', value: 10 });
  assert.equal(table.rows[0].cells[512].props.width, undefined);
});

test('a cell edge to the left of the one before it is pulled up to it: the cell is as narrow as a cell can be', () => {
  const table = tables(read(doc(String.raw`\trowd\cellx3000\cellx2000\cellx5000\pard\intbl a\cell b\cell c\cell\row\pard x\par`)))[0];
  assert.deepEqual(table.grid, [3000, 1, 1999]);
  assert.deepEqual(table.rows[0].cells.map((cell) => (cell.props.width as { value: number }).value), [3000, 1, 1999]);
  assert.equal(table.rows[0].cells.length, 3);
});

test('cells merged into the one before them: it spans their columns and is as wide as all of them', () => {
  const merged = String.raw`\trowd\clmgf\cellx1000\clmrg\cellx2500\clmrg\cellx3000\cellx4000\pard\intbl a\cell\cell\cell d\cell\row`;
  const plain = String.raw`\trowd\cellx1000\cellx2500\cellx3000\cellx4000\pard\intbl e\cell f\cell g\cell h\cell\row`;
  const table = tables(read(doc(merged + plain + String.raw`\pard x\par`)))[0];
  assert.deepEqual(table.grid, [1000, 1500, 500, 1000]);
  const [first, last] = table.rows[0].cells;
  assert.equal(table.rows[0].cells.length, 2);
  assert.equal(first.props.gridSpan, 3);
  assert.deepEqual(first.props.width, { type: 'dxa', value: 3000 });
  assert.equal(first.props.hMerge, undefined);
  assert.equal(last.props.gridSpan, undefined);
  assert.equal(table.rows[1].cells.length, 4);
});

// ---------------------------------------------------------------------------
// Pictures and shapes

test('a shape measured from the page: horizontally and vertically each by its own word', () => {
  const shape = (words: string) =>
    String.raw`{\shp{\*\shpinst\shpleft100\shptop200\shpright1100\shpbottom700` + words + String.raw`\shpwr3{\sp{\sn shapeType}{\sv 1}}}}\par`;
  const anchor = (words: string) => drawingsOf(paragraphs(read(doc(shape(words))))[0])[0].anchor!;
  assert.equal(anchor(String.raw`\shpbxcolumn\shpbypage`).vertical!.relativeFrom, 'page');
  assert.equal(anchor(String.raw`\shpbxcolumn\shpbypage`).horizontal!.relativeFrom, 'column');
  assert.equal(anchor(String.raw`\shpbxpage\shpbypara`).vertical!.relativeFrom, 'paragraph');
  assert.equal(anchor(String.raw`\shpbxpage\shpbypara`).horizontal!.relativeFrom, 'page');
  assert.equal(anchor(String.raw`\shpbxmargin\shpbymargin`).vertical!.relativeFrom, 'margin');
});

test('a picture larger than the reader keeps is left out, with a warning; the text around it stays', () => {
  const head = bytesOf(String.raw`{\rtf1 a{\pict\pngblip\picwgoal600\pichgoal300\bin67108865 `);
  const tail = bytesOf(String.raw`}b\par}`);
  const bytes = new Uint8Array(head.length + 67108865 + tail.length);
  bytes.set(head);
  bytes.set(tail, head.length + 67108865);
  const model = readRtf(bytes);
  assert.deepEqual(texts(model), ['ab']);
  assert.ok(model.warnings.some((warning) => /picture is larger than the reader keeps/.test(warning)));
  assert.equal(model.media, undefined);
  // It still takes its room.
  const drawing = drawingsOf(paragraphs(model)[0])[0];
  assert.equal(drawing.image, undefined);
  assert.deepEqual([drawing.width, drawing.height], [600 * 635, 300 * 635]);
});

// ---------------------------------------------------------------------------
// Styles and lists

const LIST =
  '{\\*\\listtable{\\list\\listsimple{\\listlevel\\levelnfc0\\levelstartat1{\\leveltext\\' + "'02\\'00.;}{\\levelnumbers\\'01;}}\\listid7}" +
  '{\\list\\listsimple{\\listlevel\\levelnfc23{\\leveltext\\' + "'02\\'00.;}{\\levelnumbers;}}\\listid8}}" +
  String.raw`{\*\listoverridetable{\listoverride\listid7\ls1}{\listoverride\listid7\listoverridecount1{\lfolevel\levelstartat7}\ls2}` +
  String.raw`{\listoverride\listid7\listoverridecount1{\lfolevel\listoverridestartat\levelstartat7}\ls3}{\listoverride\listid8\ls4}}`;

test('a paragraph without a list in a style that has one says so', () => {
  const sheet = String.raw`{\stylesheet{\s0 Normal;}{\s1\ls1 Numbered;}}`;
  const model = read(doc(String.raw`\pard\plain\s1\ls1 a\par\pard\plain\s1 b\par\pard\plain c\par`, sheet + LIST));
  const [a, b, c] = paragraphs(model);
  assert.deepEqual(model.styles.styles.get('Numbered')!.para!.numbering, { numId: 1, level: 0 });
  // The paragraph names its list, as the file does.
  assert.deepEqual(a.props.numbering, { numId: 1, level: 0 });
  assert.deepEqual(resolveParaProps(model, a.props).numbering, { numId: 1, level: 0 });
  // The style numbers it and the paragraph is not numbered: number 0 takes the list away.
  assert.deepEqual(b.props.numbering, { numId: 0 });
  // No style with a list: nothing to take away.
  assert.equal(c.props.numbering, undefined);
});

test('styles based on each other in a circle, and on a style of another kind, stand alone', () => {
  const loop = read(doc(String.raw`\pard\plain\s1\b a\par`, String.raw`{\stylesheet{\s1\sbasedon2\b One;}{\s2\sbasedon3\i Two;}{\s3\sbasedon1\fs40 Three;}{\s4\sbasedon4 Self;}}`));
  const styles = loop.styles.styles;
  // The circle is cut where it closes; what is left is a chain that ends.
  const chain: string[] = [];
  for (let id: string | undefined = 'One'; id !== undefined && chain.length < 10; id = styles.get(id)!.basedOn) chain.push(id);
  assert.deepEqual(chain, ['One', 'Two', 'Three']);
  assert.equal(styles.get('Self')!.basedOn, undefined);
  // Every style still gives what it states: the text says what its style says, so it states nothing itself.
  assert.equal(runWith(paragraphs(loop)[0], 'a').props.bold, undefined);
  assert.equal(effectiveRun(loop, paragraphs(loop)[0], 'a').bold, true);

  // A paragraph style based on a character style, and the other way round: not based on anything.
  const mixed = read(doc(String.raw`a\par`, String.raw`{\stylesheet{\s1\sbasedon10 Para;}{\*\cs10\sbasedon1\b Char;}{\s2\sbasedon1 Child;}{\*\cs11\sbasedon11\i Own;}}`));
  assert.equal(mixed.styles.styles.get('Para')!.basedOn, undefined);
  assert.equal(mixed.styles.styles.get('Char')!.basedOn, undefined);
  assert.equal(mixed.styles.styles.get('Child')!.basedOn, 'Para');
  // A character style based on itself.
  assert.equal(mixed.styles.styles.get('Own')!.basedOn, undefined);
  assert.equal(mixed.styles.styles.get('Own')!.run!.italic, true);
});

test('a character below 9 in the text of a level is a counter, except in a level of bullets', () => {
  const model = read(doc('a\\par', LIST));
  const numbered = model.numbering.abstracts.get(0)!.levels.get(0)!;
  assert.deepEqual([numbered.format, numbered.text], ['decimal', '%1.']);
  const bullet = model.numbering.abstracts.get(1)!.levels.get(0)!;
  assert.equal(bullet.format, 'bullet');
  // No counter in a bullet: the character is kept as it is.
  assert.equal(bullet.text, chars(0) + '.');
});

test('a list override restarts a level only when it says so', () => {
  const model = read(doc('a\\par', LIST));
  assert.equal(model.numbering.nums.get(1)!.overrides.size, 0);
  // A start value without "\listoverridestartat" is not an override.
  assert.equal(model.numbering.nums.get(2)!.overrides.size, 0);
  assert.deepEqual(model.numbering.nums.get(3)!.overrides.get(0), { start: 7 });
});

// ---------------------------------------------------------------------------
// Sections and paragraphs

test('"\\titlepg0" takes the first page of its own away again', () => {
  const pageOf = (model: DocModel) => model.body.section!.titlePage;
  assert.equal(pageOf(read(doc(String.raw`\sectd\titlepg a\par`))), true);
  assert.equal(pageOf(read(doc(String.raw`\sectd\titlepg1 a\par`))), true);
  assert.equal(pageOf(read(doc(String.raw`\sectd\titlepg\titlepg0 a\par`))), undefined);
  assert.equal(pageOf(read(doc(String.raw`\sectd\titlepg0 a\par`))), undefined);
});

test('space before and after a paragraph is never negative', () => {
  const model = read(doc(String.raw`\sb-100\sa-200 a\par\pard\sb100\sa200 b\par`));
  const [a, b] = paragraphs(model);
  assert.equal(a.props.spacing, undefined);
  assert.deepEqual([effectivePara(model, a).spacing!.before || 0, effectivePara(model, a).spacing!.after || 0], [0, 0]);
  assert.deepEqual([b.props.spacing!.before, b.props.spacing!.after], [100, 200]);
});

test('"\\hyphpar0" takes a paragraph out of hyphenation; "\\hyphpar" and "\\hyphpar1" put it back', () => {
  const model = read(doc(String.raw`\hyphpar0 a\par\hyphpar b\par\hyphpar0 c\par\hyphpar1 d\par`));
  assert.deepEqual(paragraphs(model).map((p) => effectivePara(model, p).suppressAutoHyphens === true), [true, false, true, false]);
});
