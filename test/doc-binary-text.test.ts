// The header and the text of a Word binary file: versions, the piece table,
// compressed and 16-bit pieces, the characters that are not text.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { readDoc, UnsupportedDocError } from '../src/formats/doc/document';
import { readFib } from '../src/formats/doc/fib';
import { readPieceTable } from '../src/formats/doc/piece-table';
import { paragraphText } from '../src/model/doc/types';
import { DocBuilder, FC, buildCfb, cat, paragraphs, plc, runsOf, simpleDoc, sprm, textOf, u16le, u32le } from './doc-binary-fixtures';

test('reads the paragraphs of a document of 16-bit text', () => {
  const model = simpleDoc('Merhaba dünya\rİkinci paragraf: ğüşıöç\r\rSon\r').read();
  assert.deepEqual(textOf(model.body.blocks), ['Merhaba dünya', 'İkinci paragraf: ğüşıöç', '', 'Son']);
  assert.deepEqual(model.warnings, []);
  assert.equal(model.body.source, '');
  assert.equal(model.body.part, 'doc:body');
  assert.ok(model.body.section, 'the last section is that of the body');
});

test('a paragraph of the default style states no style and shares its properties', () => {
  const model = simpleDoc('a\rb\r').read();
  const [first, second] = paragraphs(model.body.blocks);
  assert.equal(first.props.style, undefined);
  assert.equal(first.props, second.props);
});

test('compressed text is Windows-1252 with the special characters of 0x80 to 0x9F', () => {
  const bytes = [0x93, 0x41, 0x94, 0x85, 0x96, 0x97, 0x91, 0x92, 0x95, 0x99, 0xe9, 0xfc, 0x8c, 0x9f, 0x80, 0x0d];
  const text = String.fromCharCode(...bytes);
  const builder = new DocBuilder();
  builder.text(text, { compressed: true });
  builder.papxAll(text);
  const model = builder.read();
  assert.deepEqual(textOf(model.body.blocks), ['“A”…–—‘’•™éüŒŸ\u0080']);
});

test('a paragraph may run over compressed and 16-bit pieces', () => {
  const builder = new DocBuilder();
  builder.text('Hello ', { compressed: true });
  builder.text('dünya ş');
  builder.text(' again\rnext', { compressed: true });
  builder.text(' one\r');
  builder.papx(0, 20, 0).papx(20, 29, 0);
  const model = builder.read();
  assert.deepEqual(textOf(model.body.blocks), ['Hello dünya ş again', 'next one']);
});

test('pieces may lie in the stream in any order', () => {
  const builder = new DocBuilder();
  // The second piece of the text is written to the stream first.
  const late = builder.blob(new Uint8Array(0));
  builder.text('second\r', { compressed: true });
  builder.text('first ', { compressed: true });
  const model = builder.read({
    patch: ({ table, fcLcb }) => {
      // Swap the two pieces: positions 0..6 are "first " and 6..13 "second\r".
      const [at] = fcLcb(FC.Clx);
      const view = new DataView(table.buffer, table.byteOffset);
      const pieces = at + 5;
      view.setUint32(pieces + 4, 6, true);
      view.setUint32(pieces + 8, 13, true);
      const data = pieces + 12;
      view.setUint32(data + 2, ((late + 7) * 2) | 0x40000000, true);
      view.setUint32(data + 8 + 2, (late * 2) | 0x40000000, true);
    },
  });
  assert.deepEqual(textOf(model.body.blocks), ['first second']);
});

test('characters that are not text become items of their own', () => {
  const text = 'a\tb\u000bc\u000cd\u000ee\u001ef\u001fg h\r';
  const model = simpleDoc(text).read();
  const [paragraph] = paragraphs(model.body.blocks);
  const items = runsOf(paragraph).flatMap((run) => run.content);
  assert.deepEqual(
    items.map((item) => (item.type === 'text' ? item.text : item.type === 'break' ? 'break:' + item.kind : item.type)),
    ['a', 'tab', 'b', 'break:line', 'c', 'break:page', 'd', 'break:column', 'e', 'noBreakHyphen', 'f', 'softHyphen', 'g h'],
  );
  assert.equal(paragraphText(paragraph), 'a\tb\nc\nd\ne‑fg h');
});

test('control characters that stand for nothing are dropped, not shown', () => {
  const model = simpleDoc('a\u0001b\u0002c\u0005d\u0008e\u0010f\r').read();
  assert.deepEqual(textOf(model.body.blocks), ['abcdef']);
});

test('text without a paragraph mark at its end is still a paragraph', () => {
  const builder = new DocBuilder();
  builder.text('one\rtwo');
  builder.papx(0, 4, 0);
  assert.deepEqual(textOf(builder.read().body.blocks), ['one', 'two']);
});

test('a fast-saved file: the modifier of a piece formats its text', () => {
  const builder = new DocBuilder();
  // A Prm of the short kind: sprm number 0x55 (bold) with the operand 1.
  builder.text('bold ', { compressed: true, prm: (0x55 << 1) | (1 << 8) });
  // A Prm that refers to a property list: italic text in a centred paragraph.
  const prm = builder.prc(sprm(0x0836, 1), sprm(0x2461, 1));
  builder.text('italic\r', { prm });
  builder.text('plain\r');
  builder.papx(0, 12, 0).papx(12, 18, 0);
  const model = builder.read({ flags: 0x0004 });
  const [first, second] = paragraphs(model.body.blocks);
  const runs = runsOf(first);
  assert.deepEqual(
    runs.map((run) => [paragraphText({ ...first, children: [run] }), run.props.bold, run.props.italic]),
    [
      ['bold ', true, undefined],
      ['italic', undefined, true],
    ],
  );
  assert.equal(first.props.align, 'center');
  assert.equal(second.props.align, undefined);
});

test('the header says which table stream belongs to the file; the other one is used when it is the only one', () => {
  const model = simpleDoc('x\r').read({ tableName: '0Table' });
  assert.deepEqual(textOf(model.body.blocks), ['x']);
});

test('documents of Word 95 and earlier are refused with a clear error', () => {
  for (const options of [{ wIdent: 0xa5dc, nFib: 0x65 }, { nFib: 0x68 }, { wIdent: 0xa59b, nFib: 0x2d }]) {
    assert.throws(
      () => simpleDoc('x\r').read(options),
      (error: Error) => error instanceof UnsupportedDocError && error.name === 'UnsupportedFileError' && /Word 95/.test(error.message),
    );
  }
});

test('encrypted documents are refused with a clear error', () => {
  for (const flags of [0x0100, 0x8100]) {
    assert.throws(
      () => simpleDoc('x\r').read({ flags }),
      (error: Error) => error instanceof UnsupportedDocError && /Password-protected/.test(error.message),
    );
  }
});

test('a file without the streams or the header of a document is an error, not a crash', () => {
  assert.throws(() => readDoc(CfbFile.read(buildCfb([['Other', new Uint8Array(600)]]))), /Invalid doc: the file has no WordDocument stream/);
  assert.throws(() => readDoc(CfbFile.read(buildCfb([['WordDocument', new Uint8Array(20)]]))), /Invalid doc: the WordDocument stream is too short/);
  assert.throws(() => readDoc(CfbFile.read(buildCfb([['WordDocument', new Uint8Array(2048)]]))), /Invalid doc: the WordDocument stream does not start with a Word header/);
  const word = new Uint8Array(2048);
  word.set([0xec, 0xa5, 0xc1, 0x00]);
  assert.throws(() => readDoc(CfbFile.read(buildCfb([['WordDocument', word]]))), /Invalid doc: the file has no table stream/);
  assert.throws(() => simpleDoc('x\r').read({ noClx: true }), /Invalid doc: the piece table/);
});

test('the header gives the lengths of the parts and the places of the structures', () => {
  const builder = simpleDoc('abc\r');
  builder.table(FC.Dop, new Uint8Array(100));
  let fib: ReturnType<typeof readFib> | undefined;
  builder.read({
    patch: ({ word, fcLcb }) => {
      fib = readFib(word);
      assert.deepEqual(fib.fcLcb(FC.Dop), fcLcb(FC.Dop));
    },
  });
  assert.ok(fib);
  assert.equal(fib.nFib, 0xc1);
  assert.equal(fib.tableStream, '1Table');
  assert.equal(fib.ccpText, 4);
  assert.equal(fib.complex, false);
  assert.deepEqual(fib.fcLcb(FC.PlfLst), [0, 0]);
  assert.deepEqual(fib.fcLcb(5000), [0, 0]);
});

test('a header that states more text than there is gets what there is, with a warning', () => {
  const model = simpleDoc('abc\rdef\r').read({
    patch: ({ word }) => new DataView(word.buffer).setUint32(76, 0x7fffffff, true),
  });
  assert.deepEqual(textOf(model.body.blocks), ['abc', 'def']);
  assert.match(model.warnings.join('\n'), /more text than the file holds/);
});

// A piece table by hand: positions and pieces of [offset, compressed].
function clx(cps: number[], pieces: [number, boolean][]): Uint8Array {
  const body = plc(cps, pieces.map(([fc, compressed]) => cat(u16le(0), u32le(compressed ? (fc * 2) | 0x40000000 : fc), u16le(0))));
  return cat([2], u32le(body.length), body);
}

test('a piece table that makes no sense is an error', () => {
  const word = new Uint8Array(4096);
  const read = (bytes: Uint8Array) => readPieceTable(bytes, 0, bytes.length, word);
  assert.equal(read(clx([0, 10], [[100, true]])).text.length, 10);
  assert.throws(() => read(clx([0, 10, 5], [[100, true], [200, true]])), /out of order/);
  assert.throws(() => read(clx([3, 10], [[100, true]])), /does not start at the first character/);
  assert.throws(() => read(clx([0, 10], [[4090, true]])), /outside the WordDocument stream/);
  assert.throws(() => read(clx([0, 3000], [[100, false]])), /outside the WordDocument stream/);
  assert.throws(() => read(cat([7, 0, 0, 0, 0, 0])), /corrupt/);
  assert.throws(() => read(cat([1], u16le(500), [0, 0])), /cut short/);
  assert.throws(() => readPieceTable(new Uint8Array(10), 4, 100, word), /outside the table stream/);
});

test('pieces that reuse the same bytes over and over do not blow up the text', () => {
  const word = new Uint8Array(4096);
  const cps = [0];
  const pieces: [number, boolean][] = [];
  for (let i = 0; i < 100; i++) {
    cps.push((i + 1) * 4000);
    pieces.push([0, true]);
  }
  const bytes = clx(cps, pieces);
  assert.throws(() => readPieceTable(bytes, 0, bytes.length, word), /more text than the file holds/);
});
