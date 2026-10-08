// Damaged and hostile Word binary files: whatever the bytes are, the reader
// ends, with a document or with an Error, in bounded time and memory. Also how
// long a large document takes.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { readDoc } from '../src/formats/doc/document';
import { DocLayout } from '../src/layout/doc/layout';
import type { Block, DocModel } from '../src/model/doc/types';
import { fakeMeasurer } from './doc-layout-fixtures';
import { DocBuilder, FC, IN_TABLE, ROW_END, buildCfb, cat, defTable, fontTable, plc, sprm, stsh, sttb, tablesOf, textOf, u16le, u32le, u8s, utf16le } from './doc-binary-fixtures';

/** A document that uses most of what the reader reads: styles, a list, a table, fields, a picture, notes, headers, sections. */
function richDoc(): DocBuilder {
  const builder = new DocBuilder();
  const body =
    'Title\r' +
    'Item one\rItem two\r' +
    'Cell A\u0007Cell B\u0007\u0007Cell C\u0007Cell D\u0007\u0007' +
    'Page \u0013 PAGE \u00141\u0015 and a note\u0002 and a picture \u0001.\r' +
    '\u000c' +
    'Second section\r';
  const notes = '\u0002 The note\r\r';
  const header = 'Header text\r\rFooter \u0013 PAGE \u00141\u0015\r\r';
  const text = body + notes + header + '\r';
  builder.text(text.slice(0, 40), { compressed: true });
  builder.text(text.slice(40));
  builder.ccp.ftn = notes.length;
  builder.ccp.hdd = header.length;

  let at = 0;
  const para = (length: number, istd: number, ...sprms: Uint8Array[]) => {
    builder.papx(at, at + length, istd, ...sprms);
    at += length;
  };
  para(6, 1);
  para(9, 0, sprm(0x460b, u16le(1)), sprm(0x260a, 0));
  para(9, 0, sprm(0x460b, u16le(1)), sprm(0x260a, 0));
  for (let row = 0; row < 2; row++) {
    para(7, 0, IN_TABLE);
    para(7, 0, IN_TABLE);
    para(1, 0, IN_TABLE, ROW_END, defTable([0, 2000, 4000]), sprm(0x9407, u16le(300)));
  }
  const fieldLine = body.indexOf('Page ');
  para(body.indexOf('\u000c') - fieldLine, 0, sprm(0x2461, 1));
  para(1, 0);
  para(15, 0);
  // The paragraphs of the notes and of the header part.
  for (let i = at; i < text.length; i++) if (text.charCodeAt(i) === 13) para(i + 1 - at, 0);

  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 2 || (code >= 0x13 && code <= 0x15)) builder.chpx(i, i + 1, sprm(0x0855, 1));
  }
  builder.chpx(0, 5, sprm(0x0835, 0x81), sprm(0x4a43, u16le(40)));
  const picture = new Uint8Array(68 + 60);
  new DataView(picture.buffer).setUint32(0, picture.length, true);
  new DataView(picture.buffer).setUint16(4, 68, true);
  new DataView(picture.buffer).setUint16(6, 0x64, true);
  new DataView(picture.buffer).setUint16(28, 1000, true);
  new DataView(picture.buffer).setUint16(30, 1000, true);
  const blip = cat(u16le(0x6e00), u16le(0xf01e), u32le(17 + 8), new Uint8Array(17), [0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
  picture.set(blip, 68);
  builder.chpx(body.indexOf('\u0001'), body.indexOf('\u0001') + 1, sprm(0x0855, 1), sprm(0x6a03, u32le(builder.dataBlob(picture))));

  builder.table(FC.SttbfFfn, fontTable(['Times New Roman', 'Arial', 'Symbol']));
  builder.table(
    FC.Stshf,
    stsh([
      { sti: 0, name: 'Normal', chpx: sprm(0x4a43, u16le(24)) },
      { sti: 1, name: 'heading 1', base: 0, next: 0, papx: sprm(0xa413, u16le(240)), chpx: sprm(0x0835, 1) },
    ]),
  );
  const level = cat(u32le(1), u8s(0, 0), new Uint8Array(9), u8s(0), u32le(0), u32le(0), u8s(0, 4, 0, 0), sprm(0x845e, u16le(720)), u16le(2), utf16le('\u0000.'));
  const levels = new Uint8Array(18).fill(0xff);
  builder.table(FC.PlfLst, cat(u16le(1), u32le(77), u32le(0), levels, u8s(1, 0), level));
  builder.table(FC.PlfLfo, cat(u32le(1), u32le(77), u32le(0), u32le(0), u8s(0, 0, 0, 0), u32le(0xffffffff)));

  const sepx = (...sprms: Uint8Array[]) => {
    const list = cat(...sprms);
    return cat(u16le(0), u32le(builder.blob(cat(u16le(list.length), list))), u16le(0), u32le(0xffffffff));
  };
  const page = [sprm(0xb01f, u16le(6000)), sprm(0xb020, u16le(6000)), sprm(0xb021, u16le(600)), sprm(0xb022, u16le(600)), sprm(0x9023, u16le(600)), sprm(0x9024, u16le(600))];
  const firstEnd = body.indexOf('\u000c') + 1;
  builder.table(FC.PlcfSed, plc([0, firstEnd, body.length], [sepx(...page), sepx(...page, sprm(0x3009, 2))]));
  const noteRef = body.indexOf('\u0002');
  builder.table(FC.PlcffndRef, plc([noteRef, body.length], [u16le(1)]));
  builder.table(FC.PlcffndTxt, plc([0, notes.length - 1, notes.length]));
  // Six empty stories for the notes, then: no even header, the odd header, no even footer, the odd footer, no first ones.
  builder.table(FC.PlcfHdd, plc([0, 0, 0, 0, 0, 0, 0, 0, 13, 13, 31, 31, 31, 32]));
  builder.table(FC.SttbfBkmk, sttb(['mark']));
  builder.table(FC.PlcfBkf, plc([2, body.length], [cat(u16le(0), u16le(0))]));
  builder.table(FC.PlcfBkl, plc([20, body.length]));
  const dop = new Uint8Array(84);
  dop.set(u16le(708), 10);
  builder.table(FC.Dop, dop);
  return builder;
}

function allText(model: DocModel): string {
  const out: string[] = [];
  const visit = (blocks: Block[]) => {
    for (const line of textOf(blocks)) out.push(line);
    for (const table of tablesOf(blocks)) for (const row of table.rows) for (const cell of row.cells) visit(cell.blocks);
  };
  visit(model.body.blocks);
  return out.join('|');
}

test('the document the damage tests start from reads as meant', () => {
  const model = richDoc().read();
  assert.deepEqual(model.warnings, []);
  assert.equal(model.unknown.size, 0);
  assert.equal(allText(model), 'Title|Item one|Item two|Page 1 and a note and a picture .||Second section|Cell A|Cell B|Cell C|Cell D');
  assert.deepEqual(textOf(model.headers.get('hdr:0:default')!.blocks), ['Header text']);
  assert.deepEqual(textOf(model.footers.get('ftr:0:default')!.blocks), ['Footer 1']);
  assert.deepEqual(textOf(model.footnotes!.items.get(1)!.blocks), [' The note']);
  assert.equal(model.media!.size, 1);
  assert.equal(model.numbering.nums.size, 1);
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  assert.equal(layout.pages.length, 2);
});

/** Reads the bytes; the outcome is a model or an Error, never anything else. Returns whether it was read. */
function survives(bytes: Uint8Array, layout: boolean): boolean {
  let model: DocModel;
  try {
    model = readDoc(CfbFile.read(bytes));
  } catch (error) {
    assert.ok(error instanceof Error, 'what is thrown is an Error');
    assert.ok(!(error instanceof RangeError) && !(error instanceof TypeError), `a damaged file is reported, not crashed on: ${String(error && (error as Error).stack)}`);
    return false;
  }
  assert.equal(typeof allText(model), 'string');
  if (layout) {
    const engine = new DocLayout(model, fakeMeasurer, { maxPages: 30 });
    engine.layout();
    assert.ok(engine.pages.length > 0);
  }
  return true;
}

/** A small deterministic generator: the damage is the same on every run. */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 0x100000000;
  };
}

test('streams cut short anywhere: an error or a document, never a crash', () => {
  const builder = richDoc();
  let streams: { word: Uint8Array; table: Uint8Array; data: Uint8Array } | undefined;
  builder.build({ patch: (built) => (streams = { word: built.word.slice(), table: built.table.slice(), data: built.data.slice() }) });
  assert.ok(streams);
  const { word, table, data } = streams;
  let read = 0;
  let refused = 0;
  const attempt = (w: Uint8Array, t: Uint8Array, d: Uint8Array) => {
    if (survives(buildCfb([['WordDocument', w], ['1Table', t], ['Data', d.length ? d : new Uint8Array(1)]]), true)) read++;
    else refused++;
  };
  for (let length = 0; length < table.length; length += 7) attempt(word, table.subarray(0, length), data);
  for (let length = 0; length < word.length; length += 61) attempt(word.subarray(0, length), table, data);
  for (let length = 0; length < data.length; length += 11) attempt(word, table, data.subarray(0, length));
  assert.ok(read > 20, `some cuts leave a readable document (${read})`);
  assert.ok(refused > 20, `some do not (${refused})`);
});

test('random damage to every stream: an error or a document, in bounded time', () => {
  const builder = richDoc();
  let streams: { word: Uint8Array; table: Uint8Array; data: Uint8Array } | undefined;
  builder.build({ patch: (built) => (streams = { word: built.word.slice(), table: built.table.slice(), data: built.data.slice() }) });
  assert.ok(streams);
  const next = random(20261002);
  const started = Date.now();
  let read = 0;
  for (let round = 0; round < 1500; round++) {
    const word = streams.word.slice();
    const table = streams.table.slice();
    const data = streams.data.slice();
    const target = round % 3 === 0 ? word : round % 3 === 1 ? table : data;
    // A few bytes changed: single bits, extreme values, and runs of 0xFF.
    const changes = 1 + Math.floor(next() * 6);
    for (let i = 0; i < changes; i++) {
      // In the WordDocument stream: the header and the formatting pages more often than the text.
      let at = Math.floor(next() * target.length);
      if (target === word && next() < 0.5) at = next() < 0.5 ? Math.floor(next() * 0x400) : 0x800 + Math.floor(next() * (word.length - 0x800));
      const kind = next();
      if (kind < 0.4) target[at] ^= 1 << Math.floor(next() * 8);
      else if (kind < 0.7) target[at] = [0, 0xff, 0x7f, 0x80][Math.floor(next() * 4)];
      else target.fill(0xff, at, Math.min(target.length, at + 1 + Math.floor(next() * 8)));
    }
    if (survives(buildCfb([['WordDocument', word], ['1Table', table], ['Data', data]]), round % 5 === 0)) read++;
  }
  assert.ok(read > 500, `most damaged files are still read (${read})`);
  assert.ok(Date.now() - started < 60000, 'no file takes long');
});

test('lengths and counts of absurd size are cut down to what the file holds', () => {
  const absurd = (patch: (view: DataView, fcLcb: (index: number) => [number, number], table: DataView) => void) =>
    survives(
      richDoc().build({
        patch: ({ word, table, fcLcb }) => patch(new DataView(word.buffer, word.byteOffset), fcLcb, new DataView(table.buffer, table.byteOffset)),
      }),
      true,
    );
  // The lengths of all parts of the text.
  assert.equal(absurd((word) => { for (let at = 76; at <= 104; at += 4) word.setUint32(at, 0xfffffff0, true); }), true);
  // Every structure of the table stream said to be 4 GB long, or to start beyond the stream.
  assert.equal(absurd((word) => { for (let i = 0; i < 93; i++) if (i !== FC.Clx) word.setUint32(154 + i * 8 + 4, 0xffffffff, true); }), true);
  assert.equal(absurd((word) => { for (let i = 0; i < 93; i++) if (i !== FC.Clx) word.setUint32(154 + i * 8, 0xfffffff0, true); }), true);
  // The count of styles, of lists, of list overrides, of fonts.
  assert.equal(absurd((_word, fcLcb, table) => table.setUint16(fcLcb(FC.Stshf)[0] + 2, 0xffff, true)), true);
  assert.equal(absurd((_word, fcLcb, table) => table.setUint16(fcLcb(FC.PlfLst)[0], 0xffff, true)), true);
  assert.equal(absurd((_word, fcLcb, table) => table.setUint32(fcLcb(FC.PlfLfo)[0], 0xffffffff, true)), true);
  assert.equal(absurd((_word, fcLcb, table) => table.setUint16(fcLcb(FC.SttbfFfn)[0], 0xffff, true)), true);
  // Section properties that lie outside the stream.
  assert.equal(absurd((_word, fcLcb, table) => table.setUint32(fcLcb(FC.PlcfSed)[0] + 12 + 2, 0x7ffffff0, true)), true);
  // The piece table is the one thing the reader cannot do without.
  assert.equal(absurd((word) => word.setUint32(154 + FC.Clx * 8 + 4, 0xffffffff, true)), false);
});

test('a file of nothing but paragraph marks is refused beyond the limit of paragraphs', () => {
  const text = '\r'.repeat(500);
  const builder = new DocBuilder();
  builder.text(text, { compressed: true });
  const cfb = CfbFile.read(builder.build());
  assert.equal(readDoc(cfb).body.blocks.length, 500);
  assert.equal(readDoc(cfb, { maxParagraphs: 500 }).body.blocks.length, 500);
  assert.throws(() => readDoc(cfb, { maxParagraphs: 499 }), /Invalid doc: the document has more than 499 paragraphs/);
});

test('header stories that all claim the whole header text are not each read in full', () => {
  // 3000 sections; every header story says it starts at 0 and ends at the end of a long header text.
  const builder = new DocBuilder();
  const body = 'x\r';
  const header = 'header line\r'.repeat(4000);
  const text = body + header + '\r';
  builder.text(text, { compressed: true });
  builder.ccp.hdd = header.length;
  const sections = 3000;
  const cps: number[] = [];
  for (let i = 0; i < 6 + sections * 6 + 1; i++) cps.push(i % 2 ? header.length : 0);
  builder.table(FC.PlcfHdd, plc(cps));
  const ends = [0];
  const seds: Uint8Array[] = [];
  for (let i = 0; i < sections; i++) {
    ends.push(2);
    seds.push(cat(u16le(0), u32le(0xffffffff), u16le(0), u32le(0xffffffff)));
  }
  builder.table(FC.PlcfSed, plc(ends, seds));
  const cfb = CfbFile.read(builder.build());
  const started = performance.now();
  const model = readDoc(cfb);
  const elapsed = performance.now() - started;
  // One story gets the text; the others, which start where it ends, are empty.
  let paragraphsRead = 0;
  for (const story of [...model.headers.values(), ...model.footers.values()]) paragraphsRead += story.blocks.length;
  for (const note of model.footnotes ? model.footnotes.items.values() : []) paragraphsRead += note.blocks.length;
  assert.ok(paragraphsRead <= 4000, String(paragraphsRead));
  assert.ok(elapsed < 50000, `took ${Math.round(elapsed)} ms`);
});

test('many characters that refer to the same picture or the same shape read it once', () => {
  const builder = new DocBuilder();
  const count = 3000;
  const text = '\u0001'.repeat(count) + '\u0008'.repeat(count) + '\r';
  builder.text(text + 'box\r\r\r', { compressed: true });
  builder.ccp.txbx = 5;
  builder.papx(0, text.length, 0);
  const picture = new Uint8Array(68 + 33);
  const view = new DataView(picture.buffer);
  view.setUint32(0, picture.length, true);
  view.setUint16(4, 68, true);
  view.setUint16(6, 0x64, true);
  picture.set(cat(u16le(0x6e00), u16le(0xf01e), u32le(17 + 8), new Uint8Array(17), [0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]), 68);
  const at = builder.dataBlob(picture);
  builder.chpx(0, count, sprm(0x0855, 1), sprm(0x6a03, u32le(at)));
  builder.chpx(count, 2 * count, sprm(0x0855, 1));
  // One group of 2000 shapes that every drawing character refers to.
  const record = (version: number, type: number, ...content: Uint8Array[]) => {
    const body = cat(...content);
    return cat(u16le(version), u16le(type), u32le(body.length), body);
  };
  const shapeOf = (id: number, flags: number) =>
    record(0xf, 0xf004, record(2, 0xf00a, u32le(id), u32le(flags)), cat(u16le(3 | (1 << 4)), u16le(0xf00b), u32le(6), u16le(0x0080), u32le(1 << 16)));
  const shapes: Uint8Array[] = [shapeOf(9000, 0x01)];
  for (let i = 1; i < 2000; i++) shapes.push(shapeOf(9000 + i, 0xa00));
  // Every shape of the group shows the same text box.
  const entries: Uint8Array[] = [];
  for (let i = 0; i < 2; i++) entries.push(new Uint8Array(22));
  builder.table(FC.PlcftxbxTxt, plc([0, 4, 5], entries));
  builder.table(FC.DggInfo, cat(record(0xf, 0xf000), u8s(0), record(0xf, 0xf002, record(0xf, 0xf003, record(0xf, 0xf003, ...shapes)))));
  const cps: number[] = [];
  const spas: Uint8Array[] = [];
  for (let i = 0; i < count; i++) {
    cps.push(count + i);
    spas.push(cat(u32le(9000), u32le(0), u32le(0), u32le(100), u32le(100), u16le(0), u32le(0)));
  }
  cps.push(text.length);
  builder.table(FC.PlcSpaMom, plc(cps, spas));
  const cfb = CfbFile.read(builder.build());
  const started = performance.now();
  const model = readDoc(cfb);
  const elapsed = performance.now() - started;
  assert.equal(model.media!.size, 1);
  const paragraph = model.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph');
  const drawings = paragraph.children.flatMap((inline) => (inline.type === 'run' ? inline.content.filter((item) => item.type === 'drawing') : []));
  assert.equal(drawings.length, 2 * count);
  assert.equal(drawings[0], drawings[1]);
  // The text box that two thousand shapes of the group share is in the drawing once, and the list is shared.
  const first = drawings[count];
  const second = drawings[count + 1];
  assert.ok(first.type === 'drawing' && second.type === 'drawing');
  assert.equal(first.textBoxes!.length, 1);
  assert.deepEqual(textOf(first.textBoxes![0]), ['box']);
  assert.equal(first.textBoxes, second.textBoxes);
  assert.ok(elapsed < 30000, `took ${Math.round(elapsed)} ms`);
});

test('property lists in the Data stream that point on to each other are followed once, not multiplied', () => {
  const builder = new DocBuilder();
  const text = 'x\r';
  builder.text(text);
  // An inner list of many table sprms; an outer list that points to it hundreds of times; a paragraph that
  // points to the outer list many times. Followed without limit this is millions of sprms from a few kilobytes.
  const inner = cat(...Array.from({ length: 2000 }, () => sprm(0x3404, 1)));
  const innerAt = builder.dataBlob(cat(u16le(inner.length), inner));
  const outer = cat(sprm(0x2461, 1), ...Array.from({ length: 1000 }, () => sprm(0x646b, u32le(innerAt))));
  const outerAt = builder.dataBlob(cat(u16le(outer.length), outer));
  builder.papx(0, 2, 0, ...Array.from({ length: 60 }, () => sprm(0x646b, u32le(outerAt))));
  const cfb = CfbFile.read(builder.build());
  const started = performance.now();
  const model = readDoc(cfb);
  assert.ok(performance.now() - started < 1000);
  // What the first list says is read.
  assert.equal(model.body.blocks[0].type === 'paragraph' && model.body.blocks[0].props.align, 'center');
});

test('many sections that share their properties, and many rows that share a costly definition, are read once', () => {
  const builder = new DocBuilder();
  const rows = 4000;
  const text = 'a\u0007\u0007'.repeat(rows) + '\r';
  builder.text(text, { compressed: true });
  const widen = Array.from({ length: 50 }, () => sprm(0x7623, 0, 63, u16le(10)));
  const cell = [IN_TABLE];
  const row = [IN_TABLE, ROW_END, defTable(Array.from({ length: 64 }, (_, i) => i * 100)), ...widen];
  for (let i = 0; i < rows; i++) builder.papx(i * 3, i * 3 + 2, 0, ...cell).papx(i * 3 + 2, i * 3 + 3, 0, ...row);
  builder.papx(rows * 3, rows * 3 + 1, 0);
  const list = cat(...Array.from({ length: 10000 }, () => sprm(0xb01f, u16le(5000))));
  const at = builder.blob(cat(u16le(list.length), list));
  const sections = 5000;
  const ends = [0];
  const seds: Uint8Array[] = [];
  for (let i = 0; i < sections; i++) {
    ends.push(Math.min(text.length, i + 1));
    seds.push(cat(u16le(0), u32le(at), u16le(0), u32le(0xffffffff)));
  }
  ends[sections] = text.length;
  builder.table(FC.PlcfSed, plc(ends, seds));
  const cfb = CfbFile.read(builder.build());
  const started = performance.now();
  const model = readDoc(cfb);
  assert.ok(performance.now() - started < 3000, `took ${Math.round(performance.now() - started)} ms`);
  assert.equal(model.body.section!.page!.width, 5000);
  const tables = tablesOf(model.body.blocks);
  assert.equal(tables.reduce((sum, table) => sum + table.rows.length, 0), rows);
});

test('a document of a megabyte of text is read quickly', () => {
  const builder = new DocBuilder();
  const parts: string[] = [];
  let length = 0;
  let index = 0;
  const target = 1024 * 1024;
  while (length < target) {
    const line = `Paragraph ${index} with some text that is long enough to wrap, çğıöşü.\r`;
    builder.papx(length, length + line.length, 0, sprm(0x845e, u16le(index % 1000)));
    builder.chpx(length, length + 9, sprm(0x0835, 1));
    builder.chpx(length + 9, length + line.length - 1, sprm(0x4a43, u16le(20 + (index % 7))));
    parts.push(line);
    length += line.length;
    index++;
  }
  builder.text(parts.join(''));
  const bytes = builder.build();
  const cfb = CfbFile.read(bytes);
  // Best of three: the first run pays for compiling the reader.
  let best = Infinity;
  let model: DocModel | undefined;
  for (let run = 0; run < 3; run++) {
    const started = performance.now();
    model = readDoc(cfb);
    best = Math.min(best, performance.now() - started);
  }
  assert.ok(model);
  assert.equal(model.body.blocks.length, index);
  assert.equal(textOf(model.body.blocks.slice(index - 1))[0], parts[index - 1].slice(0, -1));
  // Paragraphs and runs that say the same share their property objects: far fewer objects than paragraphs.
  const props = new Set(model.body.blocks.map((block) => (block.type === 'paragraph' ? block.props : undefined)));
  assert.equal(props.size, 1000);
  // Measured: about 0.15 s on a laptop of 2023. The limit is generous so that a slow machine does not fail the test.
  assert.ok(best < 3000, `reading ${length} characters took ${Math.round(best)} ms`);
});
