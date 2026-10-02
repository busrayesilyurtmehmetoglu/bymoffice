// Character and paragraph formatting, lists.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { EMPTY_PARA_PROPS, EMPTY_RUN_PROPS, type Run } from '../src/model/doc/types';
import { DocHistory, indentList, patchParaProps, patchRunProps, rangeFormat, setParaFormat, setRunFormat, toggleList, toggleRunFormat, type DocRange } from '../src/ops/doc/index';
import { docPicture, open, p, paragraphAt, pos, r, range, run, shapeAt, tbl, tc, texts, tr } from './doc-ops-fixtures';

const history = (doc: Parameters<typeof texts>[0]) => new DocHistory<DocRange | undefined>(doc);

test('formatting a range cuts the runs at its ends and joins what becomes alike', async () => {
  const doc = await open(p(r('abcdef') + r('ghi', '<w:b/>')) + p('second'));
  const h = history(doc);
  run(h, setRunFormat(doc, range(pos(0, 2), pos(0, 4)), { bold: true }));
  assert.deepEqual(shapeAt(doc, 0), [':ab', 'b:cd', ':ef', 'b:ghi']);
  run(h, setRunFormat(doc, range(pos(0, 4), pos(0, 6)), { bold: true }));
  // The run that was bold before states only the plain flag, so it stays a run of its own.
  assert.deepEqual(shapeAt(doc, 0), [':ab', 'b:cdef', 'b:ghi']);
  run(h, setRunFormat(doc, range(pos(0, 6), pos(0, 9)), { bold: true }));
  assert.deepEqual(shapeAt(doc, 0), [':ab', 'b:cdefghi']);
  // Removing it gives one run again.
  run(h, setRunFormat(doc, range(pos(0, 0), pos(0, 9)), { bold: null }));
  assert.deepEqual(shapeAt(doc, 0), [':abcdefghi']);
  assert.equal((paragraphAt(doc, 0).children[0] as Run).props, EMPTY_RUN_PROPS);
  // Bold sets both the plain and the complex script flag, like Word.
  run(h, setRunFormat(doc, range(pos(0, 0), pos(0, 1)), { bold: true, italic: true, size: 28 }));
  assert.deepEqual((paragraphAt(doc, 0).children[0] as Run).props, { bold: true, boldCs: true, italic: true, italicCs: true, size: 28, sizeCs: 28 });
  // A collapsed range changes nothing; formatting that is there already neither.
  assert.equal(setRunFormat(doc, range(pos(0, 3), pos(0, 3)), { bold: true }).ops.length, 0);
  assert.equal(setRunFormat(doc, range(pos(0, 0), pos(0, 1)), { bold: true }).ops.length, 0);
});

test('the property bags of the document are never changed', async () => {
  const doc = await open(p(r('abc', '<w:i/>') + r('def', '<w:i/>')));
  const first = (paragraphAt(doc, 0).children[0] as Run).props;
  const copy = JSON.stringify(first);
  run(history(doc), setRunFormat(doc, range(pos(0, 1), pos(0, 5)), { color: 'ff0000', font: 'Mono', underline: 'double', highlight: 'yellow', vertAlign: 'superscript', style: 'Strong' }));
  assert.equal(JSON.stringify(first), copy);
  assert.deepEqual(Object.keys(EMPTY_RUN_PROPS), []);
  assert.deepEqual(Object.keys(EMPTY_PARA_PROPS), []);
  const props = (paragraphAt(doc, 0).children[1] as Run).props;
  assert.deepEqual(props.color, { rgb: 'FF0000' });
  assert.deepEqual(props.fonts, { ascii: 'Mono', hAnsi: 'Mono', eastAsia: 'Mono', cs: 'Mono' });
  assert.deepEqual(props.underline, { style: 'double' });
  assert.equal(props.style, 'Strong');
  assert.equal(props.italic, true);
  assert.deepEqual(shapeAt(doc, 0).map((s) => s.split(':')[1]), ['a', 'bcde', 'f']);
});

test('a range over several paragraphs formats their marks too, and cells of tables', async () => {
  const doc = await open(p('one') + tbl(tr(tc('a') + tc('b')), [1000, 1000]) + p('two'));
  const h = history(doc);
  run(h, setRunFormat(doc, range(pos(0, 1), pos(2, 2)), { italic: true }));
  assert.deepEqual(shapeAt(doc, 0), [':o', 'i:ne']);
  assert.deepEqual(shapeAt(doc, [1, 0, 0, 0]), ['i:a']);
  assert.deepEqual(shapeAt(doc, [1, 0, 1, 0]), ['i:b']);
  assert.deepEqual(shapeAt(doc, 2), ['i:tw', ':o']);
  assert.equal(paragraphAt(doc, 0).props.markProps!.italic, true);
  assert.equal(paragraphAt(doc, [1, 0, 1, 0]).props.markProps!.italic, true);
  assert.equal(paragraphAt(doc, 2).props.markProps, undefined, 'the mark of the last paragraph is outside the range');
  h.undo();
  assert.deepEqual(shapeAt(doc, [1, 0, 0, 0]), [':a']);
});

test('rangeFormat reports the formatting in effect, styles included', async () => {
  const doc = await open(
    p(r('plain ') + r('bold', '<w:b/>') + r('strong', '<w:rStyle w:val="Strong"/>') + r('big', '<w:sz w:val="40"/><w:color w:val="00FF00"/><w:u w:val="single"/>')) +
      p('heading', '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/>') +
      p('item', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>') +
      p('', '<w:rPr><w:i/></w:rPr>'),
  );
  const plain = rangeFormat(doc, range(pos(0, 0), pos(0, 6)));
  assert.deepEqual(plain, { bold: false, italic: false, underline: false, strike: false, font: 'Test', size: 24, color: 'auto', highlight: 'none', vertAlign: 'baseline', style: '', align: 'left', paraStyle: '', list: 'none', mixed: [] });
  assert.equal(rangeFormat(doc, range(pos(0, 6), pos(0, 16))).bold, true, 'bold from the run and bold from the character style');
  assert.equal(rangeFormat(doc, range(pos(0, 6), pos(0, 16))).style, undefined);
  const mixed = rangeFormat(doc, range(pos(0, 0), pos(0, 19)));
  assert.equal(mixed.bold, undefined);
  assert.deepEqual(mixed.mixed.sort(), ['bold', 'color', 'size', 'style', 'underline']);
  const big = rangeFormat(doc, range(pos(0, 16), pos(0, 19)));
  assert.equal(big.size, 40);
  assert.equal(big.color, '00FF00');
  assert.equal(big.underline, true);
  // The heading is bold through its paragraph style.
  const heading = rangeFormat(doc, range(pos(1, 0), pos(1, 7)));
  assert.equal(heading.bold, true);
  assert.equal(heading.size, 32);
  assert.equal(heading.align, 'center');
  assert.equal(heading.paraStyle, 'Heading1');
  const item = rangeFormat(doc, pos(2, 1));
  assert.equal(item.list, 'number');
  assert.equal(item.level, 1);
  // A caret: what typed text would take. At a border, the run to the left; in an empty paragraph, the mark.
  assert.equal(rangeFormat(doc, pos(0, 6)).bold, false);
  assert.equal(rangeFormat(doc, pos(0, 10)).bold, true);
  assert.equal(rangeFormat(doc, pos(3, 0)).italic, true);
  assert.deepEqual(rangeFormat(doc, range(pos(1, 0), pos(2, 4))).mixed.sort(), ['align', 'bold', 'level', 'list', 'paraStyle', 'size']);
});

test('toggling turns a property on unless all of the range has it; off also against a style', async () => {
  const doc = await open(p(r('plain ') + r('bold', '<w:b/>')) + p('heading', '<w:pStyle w:val="Heading1"/>'));
  const h = history(doc);
  run(h, toggleRunFormat(doc, range(pos(0, 0), pos(0, 10)), 'bold'));
  assert.deepEqual(shapeAt(doc, 0), ['b:plain bold']);
  run(h, toggleRunFormat(doc, range(pos(0, 0), pos(0, 10)), 'bold'));
  assert.deepEqual(shapeAt(doc, 0), [':plain bold']);
  // The style makes the heading bold: removing the property does not help, so it is switched off on the text.
  run(h, toggleRunFormat(doc, range(pos(1, 0), pos(1, 4)), 'bold'));
  assert.deepEqual(shapeAt(doc, 1), ['B:head', ':ing']);
  assert.equal(rangeFormat(doc, range(pos(1, 0), pos(1, 4))).bold, false);
  run(h, toggleRunFormat(doc, range(pos(1, 0), pos(1, 4)), 'bold'));
  assert.equal(rangeFormat(doc, range(pos(1, 0), pos(1, 7))).bold, true);
  run(h, toggleRunFormat(doc, range(pos(0, 0), pos(0, 5)), 'underline'));
  assert.deepEqual(shapeAt(doc, 0), ['u:plain', ': bold']);
  run(h, toggleRunFormat(doc, range(pos(0, 0), pos(0, 5)), 'underline'));
  assert.deepEqual(shapeAt(doc, 0), [':plain bold']);
});

test('invalid formats are refused before anything changes', async () => {
  const doc = await open(p('text'));
  const before = docPicture(doc);
  const all = range(pos(0, 0), pos(0, 4));
  assert.throws(() => setRunFormat(doc, all, { size: 0 }), RangeError);
  assert.throws(() => setRunFormat(doc, all, { size: 1.5 }), RangeError);
  assert.throws(() => setRunFormat(doc, all, { color: 'red' }), RangeError);
  assert.throws(() => setRunFormat(doc, all, { font: '' }), RangeError);
  assert.throws(() => setRunFormat(doc, all, { underline: '<x>' }), RangeError);
  assert.throws(() => setRunFormat(doc, all, { vertAlign: 'up' as never }), RangeError);
  assert.throws(() => setRunFormat(doc, range(pos(0, 0), pos(0, 0)), { size: -1 }), RangeError);
  assert.throws(() => setParaFormat(doc, all, { align: 'a b' }), RangeError);
  assert.throws(() => setParaFormat(doc, all, { numbering: { numId: -1 } }), RangeError);
  assert.throws(() => setParaFormat(doc, all, { numbering: { numId: 1, level: 9 } }), RangeError);
  assert.throws(() => setParaFormat(doc, all, { indent: { left: NaN } }), RangeError);
  assert.throws(() => setParaFormat(doc, all, { outlineLevel: 12 }), RangeError);
  assert.throws(() => toggleRunFormat(doc, all, 'size' as never), RangeError);
  assert.throws(() => toggleList(doc, all, 'other' as never), RangeError);
  assert.throws(() => indentList(doc, all, 0.5), RangeError);
  assert.equal(docPicture(doc), before);
});

test('paragraph formatting applies to every paragraph the range touches', async () => {
  const doc = await open(p('one', '<w:ind w:left="100" w:hanging="50"/>') + p('two') + p('three'));
  const h = history(doc);
  const third = doc.body.blocks[2];
  run(h, setParaFormat(doc, range(pos(0, 2), pos(1, 0)), { align: 'center', spacing: { before: 120, line: 360, lineRule: 'auto' }, indent: { firstLine: 200 }, style: 'Heading1' }));
  assert.deepEqual(paragraphAt(doc, 0).props, { style: 'Heading1', align: 'center', spacing: { before: 120, line: 360, lineRule: 'auto' }, indent: { left: 100, firstLine: 200 } });
  assert.deepEqual(paragraphAt(doc, 1).props, { style: 'Heading1', align: 'center', spacing: { before: 120, line: 360, lineRule: 'auto' }, indent: { firstLine: 200 } });
  assert.equal(doc.body.blocks[2], third, 'a paragraph outside the range is not touched');
  run(h, setParaFormat(doc, pos(0, 0), { align: null, spacing: { before: null }, indent: null, style: null }));
  assert.deepEqual(paragraphAt(doc, 0).props, { spacing: { line: 360, lineRule: 'auto' } });
  assert.equal(setParaFormat(doc, pos(2, 0), { align: null }).ops.length, 0);
  assert.equal(patchParaProps(EMPTY_PARA_PROPS, {}), EMPTY_PARA_PROPS);
  assert.equal(patchRunProps(EMPTY_RUN_PROPS, { bold: null }), EMPTY_RUN_PROPS);
});

test('toggleList makes a list definition when the document has none to continue, and undoing takes it away', async () => {
  const doc = await open({ body: p('one') + p('two') + p('three') + tbl(tr(tc('cell')), [1000]) + p('four'), numbering: '' });
  const h = history(doc);
  assert.equal(doc.numbering.nums.size, 0);
  run(h, toggleList(doc, range(pos(0, 0), pos(1, 1)), 'bullet'));
  assert.equal(doc.numbering.nums.size, 1);
  assert.equal(doc.numbering.abstracts.size, 1);
  assert.equal(doc.numbering.modified, true);
  const numId = paragraphAt(doc, 0).props.numbering!.numId!;
  assert.ok(numId >= 1);
  assert.deepEqual(paragraphAt(doc, 1).props.numbering, { numId, level: 0 });
  const abstract = doc.numbering.abstracts.get(doc.numbering.nums.get(numId)!.abstractId)!;
  assert.equal(abstract.levels.size, 9);
  assert.equal(abstract.levels.get(0)!.format, 'bullet');
  assert.deepEqual(abstract.levels.get(2)!.para, { indent: { left: 2160, hanging: 360 } });
  assert.equal(rangeFormat(doc, pos(0, 0)).list, 'bullet');
  // The next paragraph continues the list of its neighbour instead of starting another.
  run(h, toggleList(doc, pos(2, 0), 'bullet'));
  assert.equal(doc.numbering.nums.size, 1);
  assert.equal(paragraphAt(doc, 2).props.numbering!.numId, numId);
  // A numbered list is another list; the paragraph in the cell does not see the list before the table.
  run(h, toggleList(doc, pos([3, 0, 0, 0], 0), 'number'));
  assert.equal(doc.numbering.nums.size, 2);
  assert.equal(rangeFormat(doc, pos([3, 0, 0, 0], 0)).list, 'number');
  assert.equal(doc.numbering.abstracts.get(doc.numbering.nums.get(paragraphAt(doc, [3, 0, 0, 0]).props.numbering!.numId!)!.abstractId)!.levels.get(1)!.text, '%2.');
  // All of the range is in the list already: out of it.
  run(h, toggleList(doc, range(pos(0, 0), pos(2, 0)), 'bullet'));
  assert.equal(paragraphAt(doc, 0).props.numbering, undefined);
  assert.equal(paragraphAt(doc, 2).props.numbering, undefined);
  while (h.undo());
  assert.equal(doc.numbering.nums.size, 0);
  assert.equal(doc.numbering.abstracts.size, 0);
  assert.equal('modified' in doc.numbering, false);
});

test('a bulleted item becomes a numbered one at the depth it has; levels stay within 0 and 8', async () => {
  const doc = await open(p('a', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>') + p('b', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>') + p('c'));
  const h = history(doc);
  run(h, toggleList(doc, pos(0, 0), 'number'));
  assert.deepEqual(paragraphAt(doc, 0).props.numbering, { numId: 2, level: 1 }, 'continues the numbered list after it');
  run(h, indentList(doc, range(pos(0, 0), pos(2, 0)), 1));
  assert.equal(paragraphAt(doc, 0).props.numbering!.level, 2);
  assert.equal(paragraphAt(doc, 1).props.numbering!.level, 1);
  assert.deepEqual(paragraphAt(doc, 2).props.indent, { left: 720 }, 'outside a list the paragraph is indented');
  for (let i = 0; i < 10; i++) run(h, indentList(doc, range(pos(0, 0), pos(2, 0)), 1));
  assert.equal(paragraphAt(doc, 0).props.numbering!.level, 8);
  for (let i = 0; i < 12; i++) {
    const edit = indentList(doc, range(pos(0, 0), pos(2, 0)), -1);
    if (edit.ops.length) run(h, edit);
  }
  assert.equal(paragraphAt(doc, 0).props.numbering!.level, 0);
  assert.equal(paragraphAt(doc, 2).props.indent, undefined);
  assert.equal(indentList(doc, range(pos(0, 0), pos(2, 0)), -1).ops.length, 0);
});
