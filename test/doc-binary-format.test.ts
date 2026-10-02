// Character and paragraph formatting of a Word binary file: formatted disk
// pages, sprm lists, the style sheet and the fonts.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { eachSprm } from '../src/formats/doc/sprm';
import { resolveParaProps, resolveRunProps, styleChain } from '../src/formats/docx/resolve';
import { findWrapper, paragraphText, type RunProps } from '../src/model/doc/types';
import { DocBuilder, FC, brc80, cat, defTable, fontTable, paragraphs, runsOf, sprm, stsh, sttb, textOf, u16le, u32le, u8s } from './doc-binary-fixtures';

/** A document of one paragraph whose first `length` characters have the sprms. */
function withRun(text: string, length: number, ...sprms: Uint8Array[]): DocBuilder {
  const builder = new DocBuilder();
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(0, length, ...sprms);
  return builder;
}

const firstRun = (builder: DocBuilder): RunProps => runsOf(paragraphs(builder.read().body.blocks)[0])[0].props;

test('the operand size of a sprm follows from its code, for every size class', () => {
  // Unknown character sprms (kind 2) of each of the eight size classes, then one that is known.
  const unknown = (spra: number) => (spra << 13) | 0x0800 | 0x01f0;
  const list = cat(
    sprm(unknown(0), 1),
    sprm(unknown(1), 2),
    sprm(unknown(2), 3, 4),
    sprm(unknown(3), 5, 6, 7, 8),
    sprm(unknown(4), 9, 10),
    sprm(unknown(5), 11, 12),
    sprm(unknown(6), 3, 13, 14, 15),
    sprm(unknown(7), 16, 17, 18),
    sprm(0x0835, 1),
  );
  const seen: [number, number][] = [];
  assert.equal(eachSprm(list, 0, list.length, (op, _at, length) => seen.push([op >> 13, length])), true);
  assert.deepEqual(seen, [[0, 1], [1, 1], [2, 2], [3, 4], [4, 2], [5, 2], [6, 3], [7, 3], [0, 1]]);

  const builder = withRun('bold\r', 4, list);
  const model = builder.read();
  assert.equal(runsOf(paragraphs(model.body.blocks)[0])[0].props.bold, true);
  // The unknown ones were passed over and counted, by code.
  assert.equal(model.unknown.size, 8);
  assert.equal(model.unknown.get('sprm/0x09F0'), 1);
  assert.equal(model.unknown.get('sprm/0xE9F0'), 1);
});

test('the sprms with a size of their own: table definition and long tab lists', () => {
  // sprmTDefTable has a size of two bytes that counts one more than follows.
  const table = defTable([0, 1000, 2000]);
  // sprmPChgTabs with the size 255: the length follows from the counts inside (2 deleted, 1 added).
  const tabs = cat(u16le(0xc615), u8s(255, 2), u16le(100), u16le(200), u16le(0), u16le(0), u8s(1), u16le(720), u8s(1));
  const list = cat(table, tabs, sprm(0x2461, 2));
  const seen: [number, number][] = [];
  assert.equal(eachSprm(list, 0, list.length, (op, _at, length) => seen.push([op, length])), true);
  assert.deepEqual(seen, [[0xd608, 7], [0xc615, 13], [0x2461, 1]]);
});

test('a sprm list that ends inside a sprm stops there, and says so', () => {
  const list = cat(sprm(0x0835, 1), u16le(0x6a03), u8s(1, 2));
  const seen: number[] = [];
  assert.equal(eachSprm(list, 0, list.length, (op) => seen.push(op)), false);
  assert.deepEqual(seen, [0x0835]);
  // A size byte that reaches past the end.
  const long = cat(u16le(0xc60d), u8s(200, 1, 2));
  assert.equal(eachSprm(long, 0, long.length, () => assert.fail('nothing to hand out')), false);
});

test('character properties', () => {
  const builder = withRun(
    'formatted text\r',
    9,
    sprm(0x0835, 1), // bold
    sprm(0x0836, 1), // italic
    sprm(0x0837, 1), // strike
    sprm(0x083a, 1), // small caps
    sprm(0x083b, 0), // caps off
    sprm(0x083c, 1), // hidden
    sprm(0x2a53, 1), // double strike
    sprm(0x4a43, u16le(28)), // 14 pt
    sprm(0x4a61, u16le(30)),
    sprm(0x4a4f, u16le(1)),
    sprm(0x4a50, u16le(2)),
    sprm(0x4a51, u16le(1)),
    sprm(0x4a5e, u16le(0)),
    sprm(0x2a42, 6), // red
    sprm(0x2a3e, 3), // double underline
    sprm(0x6877, u8s(0, 0x80, 0xff, 0)), // underline colour
    sprm(0x2a0c, 7), // yellow highlight
    sprm(0x2a48, 1), // superscript
    sprm(0x4845, u16le(0xfffa)), // lowered by 3 pt
    sprm(0x8840, u16le(20)), // spaced by 1 pt
    sprm(0x4852, u16le(150)),
    sprm(0x484b, u16le(24)),
    sprm(0x486d, u16le(0x041f)),
    sprm(0x4866, u16le((3 << 10) | (8 << 5) | 2)), // 10% blue on white
    sprm(0x6865, u32le(brc80(4, 1, 2, 1))),
    sprm(0x0875, 1),
    sprm(0x2859, 3),
  );
  builder.table(FC.SttbfFfn, fontTable(['Times New Roman', 'Arial', 'MS Mincho']));
  assert.deepEqual(firstRun(builder), {
    bold: true,
    italic: true,
    strike: true,
    smallCaps: true,
    caps: false,
    vanish: true,
    doubleStrike: true,
    size: 28,
    sizeCs: 30,
    fonts: { ascii: 'Arial', eastAsia: 'MS Mincho', hAnsi: 'Arial', cs: 'Times New Roman' },
    color: { rgb: 'FF0000' },
    underline: { style: 'double', color: { rgb: '0080FF' } },
    highlight: 'yellow',
    vertAlign: 'superscript',
    position: -6,
    spacing: 20,
    scale: 150,
    kerning: 24,
    lang: { val: 'tr-TR' },
    shading: { pattern: 'pct10', color: { rgb: '0000FF' }, fill: { rgb: 'FFFFFF' } },
    border: { style: 'single', size: 4, space: 1, color: { rgb: '0000FF' } },
    noProof: true,
    effect: 'sparkle',
  });
});

test('a full colour wins over the colour index before it, and "automatic" is said as such', () => {
  assert.deepEqual(firstRun(withRun('x\r', 1, sprm(0x2a42, 2), sprm(0x6870, u8s(0x12, 0x34, 0x56, 0)))).color, { rgb: '123456' });
  assert.deepEqual(firstRun(withRun('x\r', 1, sprm(0x6870, u8s(0, 0, 0, 0xff)))).color, { auto: true });
  assert.deepEqual(firstRun(withRun('x\r', 1, sprm(0x2a42, 0))).color, { auto: true });
});

test('runs that say the same share one property object', () => {
  const builder = new DocBuilder();
  const text = 'aa bb cc\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(0, 2, sprm(0x0835, 1)).chpx(3, 5, sprm(0x0836, 1)).chpx(6, 8, sprm(0x0835, 1));
  const runs = runsOf(paragraphs(builder.read().body.blocks)[0]);
  assert.deepEqual(runs.map((run) => paragraphText({ type: 'paragraph', props: {}, children: [run] })), ['aa', ' ', 'bb', ' ', 'cc']);
  assert.equal(runs[0].props, runs[4].props);
  assert.equal(runs[1].props, runs[3].props);
  assert.notEqual(runs[0].props, runs[2].props);
});

test('paragraph properties', () => {
  const builder = new DocBuilder();
  const text = 'paragraph\r';
  builder.text(text);
  builder.papx(
    0,
    text.length,
    0,
    sprm(0x2461, 3), // justified
    sprm(0x845e, u16le(720)),
    sprm(0x845d, u16le(360)),
    sprm(0x8460, u16le(0x10000 - 360)), // hanging
    sprm(0xa413, u16le(120)),
    sprm(0xa414, u16le(240)),
    sprm(0x6412, u16le(360), u16le(1)), // one and a half lines
    sprm(0x2405, 1),
    sprm(0x2406, 1),
    sprm(0x2407, 1),
    sprm(0x2431, 0),
    sprm(0x2640, 2),
    sprm(0x260a, 1),
    sprm(0x460b, u16le(3)),
    sprm(0x6424, u32le(brc80(8, 1, 1, 4))),
    sprm(0x6426, u32le(brc80(12, 3))),
    sprm(0x442d, u16le((1 << 10) | (16 << 5) | 1)),
    // Tabs: none deleted; a centred one at 2 cm and a right one with dots at 10 cm.
    sprm(0xc60d, 8, 0, 2, u16le(1134), u16le(5670), 1, 2 | (1 << 3)),
    sprm(0x246d, 1),
  );
  builder.chpx(text.length - 1, text.length, sprm(0x0835, 1));
  const [paragraph] = paragraphs(builder.read().body.blocks);
  assert.deepEqual(paragraph.props, {
    align: 'both',
    indent: { left: 720, right: 360, hanging: 360 },
    spacing: { before: 120, after: 240, line: 360, lineRule: 'auto' },
    keepLines: true,
    keepNext: true,
    pageBreakBefore: true,
    widowControl: false,
    outlineLevel: 2,
    numbering: { level: 1, numId: 3 },
    borders: { top: { style: 'single', size: 8, space: 4, color: { rgb: '000000' } }, bottom: { style: 'double', size: 12, space: 0, color: { auto: true } } },
    shading: { pattern: 'solid', color: { rgb: '000000' }, fill: { rgb: 'C0C0C0' } },
    tabs: [
      { pos: 1134, align: 'center' },
      { pos: 5670, align: 'right', leader: 'dot' },
    ],
    contextualSpacing: true,
    // The formatting of the paragraph mark.
    markProps: { bold: true },
  });
});

test('line spacing: a multiple of a line, at least, or exactly', () => {
  const read = (...operand: Uint8Array[]) => {
    const builder = new DocBuilder();
    builder.text('x\r');
    builder.papx(0, 2, 0, sprm(0x6412, ...operand));
    return paragraphs(builder.read().body.blocks)[0].props.spacing;
  };
  assert.deepEqual(read(u16le(480), u16le(1)), { line: 480, lineRule: 'auto' });
  assert.deepEqual(read(u16le(300), u16le(0)), { line: 300, lineRule: 'atLeast' });
  assert.deepEqual(read(u16le(0x10000 - 300), u16le(0)), { line: 300, lineRule: 'exact' });
});

test('tab lists delete and add; the long form with tolerances reads the same', () => {
  const builder = new DocBuilder();
  builder.text('a\rb\r');
  builder.papx(0, 2, 0, sprm(0xc60d, 7, 1, u16le(720), 1, u16le(1440), 3));
  builder.papx(2, 4, 0, sprm(0xc615, 9, 1, u16le(720), u16le(25), 1, u16le(1440), 3));
  const [first, second] = paragraphs(builder.read().body.blocks);
  const expected = [
    { pos: 720, align: 'clear' },
    { pos: 1440, align: 'decimal' },
  ];
  assert.deepEqual(first.props.tabs, expected);
  assert.deepEqual(second.props.tabs, expected);
});

test('a tab list with more entries than are taken is still read to its end', () => {
  const builder = new DocBuilder();
  builder.text('a\r');
  const deleted = Array.from({ length: 70 }, (_, i) => u16le(100 + i));
  builder.papx(0, 2, 0, sprm(0xc60d, 1 + 140 + 1 + 3, 70, ...deleted, 1, u16le(5000), 2));
  const tabs = paragraphs(builder.read().body.blocks)[0].props.tabs!;
  assert.equal(tabs.length, 65);
  assert.deepEqual(tabs[64], { pos: 5000, align: 'right' });
});

test('a property list too long for its page lies in the Data stream', () => {
  const builder = new DocBuilder();
  builder.text('huge\rplain\r');
  builder.dataBlob(new Uint8Array(10));
  const list = cat(sprm(0x2461, 1), sprm(0x845e, u16le(1000)));
  const at = builder.dataBlob(cat(u16le(list.length), list));
  // What stands beside the pointer is replaced by the list it points to.
  builder.papx(0, 5, 0, sprm(0x2461, 2), sprm(0x6646, u32le(at)));
  builder.papx(5, 11, 0);
  const [huge, plain] = paragraphs(builder.read().body.blocks);
  assert.deepEqual(huge.props, { align: 'center', indent: { left: 1000 } });
  assert.deepEqual(plain.props, {});

  const broken = new DocBuilder();
  broken.text('huge\r');
  broken.papx(0, 5, 0, sprm(0x2461, 2), sprm(0x6646, u32le(123456)));
  const model = broken.read();
  assert.deepEqual(textOf(model.body.blocks), ['huge']);
  assert.match(model.warnings.join('\n'), /outside the Data stream/);
});

test('properties are found for runs spread over many pages', () => {
  const builder = new DocBuilder();
  let text = '';
  const count = 600;
  for (let i = 0; i < count; i++) {
    const from = text.length;
    text += `p${i} tail\r`;
    builder.papx(from, text.length, 0, sprm(0x845e, u16le(i)));
    builder.chpx(from, from + String(i).length + 1, sprm(0x4a43, u16le(20 + (i % 50))));
  }
  builder.text(text, { compressed: true });
  const list = paragraphs(builder.read().body.blocks);
  assert.equal(list.length, count);
  for (let i = 0; i < count; i += 37) {
    assert.equal(list[i].props.indent!.left, i);
    const runs = runsOf(list[i]);
    assert.equal(runs[0].props.size, 20 + (i % 50));
    assert.equal(paragraphText(list[i]), `p${i} tail`);
    assert.equal(runs[1].props.size, undefined);
  }
});

test('a page whose first offset lies before the stretch the bin table gives it still formats its first run', () => {
  // Some writers start the offsets of every page at a round number before the text the page is for. The bin
  // table says where the page begins to count (MS-DOC 2.4.6.2): the first run reaches back no further.
  const builder = new DocBuilder();
  const count = 150;
  let text = '';
  for (let i = 0; i < count; i++) {
    builder.chpx(text.length, text.length + 2, sprm(0x4a43, u16le(20 + i)));
    text += 'ab';
  }
  text += '\r';
  builder.text(text, { compressed: true });
  builder.papx(0, text.length, 0);
  let patched = 0;
  const model = builder.read({
    patch: ({ word, table, fcLcb }) => {
      const [at, size] = fcLcb(FC.PlcfBteChpx);
      const pages = (size - 4) / 8;
      assert.ok(pages >= 3, 'the runs take several pages');
      const tableView = new DataView(table.buffer, table.byteOffset);
      const wordView = new DataView(word.buffer, word.byteOffset);
      for (let i = 1; i < pages; i++) {
        const page = tableView.getUint32(at + (pages + 1) * 4 + i * 4, true) * 512;
        assert.equal(wordView.getUint32(page, true), tableView.getUint32(at + i * 4, true));
        wordView.setUint32(page, 0x800, true);
        patched++;
      }
    },
  });
  assert.ok(patched >= 2);
  const runs = runsOf(paragraphs(model.body.blocks)[0]);
  assert.deepEqual(runs.map((run) => run.props.size), Array.from({ length: count }, (_, i) => 20 + i));
});

test('the character formatting of the paragraph style stays under a run that names a character style', () => {
  // A run in a paragraph of a bold style names the default character style and a size, and says nothing about
  // bold: it is bold. (Text that is not bold in such a paragraph says so itself.)
  const builder = styled();
  const text = 'bold by style, not bold\r';
  builder.text(text);
  builder.papx(0, text.length, 1);
  builder.chpx(0, 14, sprm(0x4a30, u16le(4)), sprm(0x4a43, u16le(22)));
  builder.chpx(14, text.length - 1, sprm(0x4a30, u16le(4)), sprm(0x0835, 0), sprm(0x4a43, u16le(22)));
  const model = builder.read();
  const [paragraph] = paragraphs(model.body.blocks);
  const [first, second] = runsOf(paragraph);
  assert.deepEqual(first.props, { style: 'DefaultParagraphFont', size: 22 });
  assert.equal(resolveRunProps(model, paragraph.props, first.props).bold, true);
  assert.equal(resolveRunProps(model, paragraph.props, second.props).bold, false);
});

const STYLES = [
  { sti: 0, name: 'Normal', chpx: cat(sprm(0x4a43, u16le(24)), sprm(0x4a4f, u16le(1))) },
  { sti: 1, name: 'heading 1', base: 0, next: 0, papx: cat(sprm(0x2461, 1), sprm(0xa413, u16le(240))), chpx: cat(sprm(0x0835, 1), sprm(0x4a43, u16le(32))) },
  { name: 'My Style', base: 1, chpx: sprm(0x0836, 1) },
  undefined,
  { sti: 65, kind: 2, name: 'Default Paragraph Font' },
  { kind: 2, name: 'Strong Text', base: 4, chpx: sprm(0x0835, 1) },
  { sti: 105, kind: 3, name: 'Normal Table', tapx: sprm(0xd634, 6, 0, 1, 2 | 8, 3, u16le(108)) },
];

function styled(): DocBuilder {
  const builder = new DocBuilder();
  builder.table(FC.Stshf, stsh(STYLES, [1, 2, 1]));
  builder.table(FC.SttbfFfn, fontTable(['Times New Roman', 'Arial', 'MS Mincho']));
  return builder;
}

test('the style sheet: ids, names, relations and formatting', () => {
  const builder = styled();
  builder.text('x\r');
  builder.papx(0, 2, 0);
  const { styles } = builder.read();
  assert.deepEqual([...styles.styles.keys()], ['Normal', 'Heading1', 'MyStyle', 'DefaultParagraphFont', 'StrongText', 'TableNormal']);
  assert.deepEqual(styles.defaultIds, { paragraph: 'Normal', character: 'DefaultParagraphFont', table: 'TableNormal' });
  const heading = styles.styles.get('Heading1')!;
  assert.equal(heading.name, 'heading 1');
  assert.equal(heading.type, 'paragraph');
  assert.equal(heading.basedOn, 'Normal');
  assert.equal(heading.next, 'Normal');
  assert.deepEqual(heading.para, { align: 'center', spacing: { before: 240 } });
  assert.deepEqual(heading.run, { bold: true, size: 32 });
  const custom = styles.styles.get('MyStyle')!;
  assert.equal(custom.custom, true);
  assert.equal(custom.basedOn, 'Heading1');
  assert.equal(styles.styles.get('StrongText')!.type, 'character');
  assert.deepEqual(styles.styles.get('TableNormal')!.table, { cellMargins: { left: { type: 'dxa', value: 108 }, right: { type: 'dxa', value: 108 } } });
  // What a docx file says in its document defaults: the default fonts of the file, 10 point.
  assert.deepEqual(styles.defaults.run, { size: 20, sizeCs: 20, fonts: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'MS Mincho' } });
});

test('the resolvers written for docx give effective values through the style chain', () => {
  const builder = styled();
  const text = 'Title\rbody strong\r';
  builder.text(text);
  builder.papx(0, 6, 2, sprm(0x845e, u16le(300)));
  builder.papx(6, text.length, 0);
  builder.chpx(11, 17, sprm(0x4a30, u16le(5)));
  const model = builder.read();
  const [title, body] = paragraphs(model.body.blocks);
  assert.equal(title.props.style, 'MyStyle');
  assert.deepEqual(styleChain(model.styles, 'MyStyle', 'paragraph').map((style) => style.id), ['Normal', 'Heading1', 'MyStyle']);
  const para = resolveParaProps(model, title.props);
  assert.equal(para.align, 'center');
  assert.deepEqual(para.indent, { left: 300 });
  assert.equal(para.spacing!.before, 240);
  const run = resolveRunProps(model, title.props, runsOf(title)[0].props);
  assert.equal(run.bold, true);
  assert.equal(run.italic, true);
  assert.equal(run.size, 32);
  assert.equal(run.fonts!.ascii, 'Arial');
  // A character style on a run of a paragraph of the default style.
  const strong = runsOf(body)[1];
  assert.equal(strong.props.style, 'StrongText');
  assert.equal(resolveRunProps(model, body.props, strong.props).bold, true);
  assert.equal(resolveRunProps(model, body.props, runsOf(body)[0].props).bold, undefined);
  assert.equal(resolveRunProps(model, body.props, runsOf(body)[0].props).size, 24);
});

test('toggles: on, off, as the style has it, the opposite of the style', () => {
  const builder = styled();
  const text = 'aabbccdd\reeff\r';
  builder.text(text);
  // A paragraph of the bold heading style, then one of the default style.
  builder.papx(0, 9, 1).papx(9, text.length, 0);
  builder.chpx(0, 2, sprm(0x0835, 0)).chpx(2, 4, sprm(0x0835, 1)).chpx(4, 6, sprm(0x0835, 0x80)).chpx(6, 8, sprm(0x0835, 0x81));
  builder.chpx(9, 11, sprm(0x0835, 0x81)).chpx(11, 13, sprm(0x0835, 0x80), sprm(0x0836, 0x81));
  const model = builder.read();
  const [heading, plain] = paragraphs(model.body.blocks);
  const effective = (paragraph: typeof heading) => runsOf(paragraph).map((run) => [run.props.bold, resolveRunProps(model, paragraph.props, run.props).bold === true]);
  assert.deepEqual(effective(heading), [
    [false, false],
    [true, true],
    [undefined, true],
    [false, false],
  ]);
  assert.deepEqual(effective(plain), [
    [true, true],
    [undefined, false],
  ]);
  assert.equal(runsOf(plain)[1].props.italic, true);
});

test('a toggle of a style may be the opposite of the style it is based on', () => {
  const builder = new DocBuilder();
  builder.table(
    FC.Stshf,
    stsh([
      { sti: 0, name: 'Normal' },
      { name: 'Bold', base: 0, chpx: sprm(0x0835, 1) },
      { name: 'Not Bold', base: 1, chpx: sprm(0x0835, 0x81) },
      { name: 'Bold Again', base: 2, chpx: sprm(0x0835, 0x81) },
      { name: 'Same', base: 1, chpx: sprm(0x0835, 0x80) },
    ]),
  );
  builder.text('x\r');
  builder.papx(0, 2, 3);
  const model = builder.read();
  const run = (id: string) => model.styles.styles.get(id)!.run;
  assert.deepEqual(run('NotBold'), { bold: false });
  assert.deepEqual(run('BoldAgain'), { bold: true });
  assert.equal(run('Same'), undefined);
  assert.equal(resolveRunProps(model, { style: 'BoldAgain' }, {}).bold, true);
});

test('styles that are based on each other in a circle end the chain instead of the reader', () => {
  const builder = new DocBuilder();
  builder.table(
    FC.Stshf,
    stsh([
      { sti: 0, name: 'Normal', base: 2, chpx: sprm(0x0835, 0x81) },
      { name: 'A', base: 2, chpx: sprm(0x0835, 0x81) },
      { name: 'B', base: 1, chpx: sprm(0x0835, 0x81) },
      { name: 'Self', base: 3, papx: sprm(0x2461, 1) },
    ]),
  );
  builder.text('x\ry\r');
  builder.papx(0, 2, 1).papx(2, 4, 3);
  builder.chpx(0, 1, sprm(0x0836, 0x81));
  const model = builder.read();
  const [first, second] = paragraphs(model.body.blocks);
  assert.equal(model.styles.styles.get('Self')!.basedOn, undefined);
  assert.equal(resolveParaProps(model, second.props).align, 'center');
  assert.equal(typeof resolveRunProps(model, first.props, runsOf(first)[0].props).bold, 'boolean');
  assert.equal(runsOf(first)[0].props.italic, true);
});

test('styles with the same name, or a name without letters, still get ids of their own', () => {
  const builder = new DocBuilder();
  builder.table(FC.Stshf, stsh([{ sti: 0, name: 'Normal' }, { name: 'Twin' }, { name: 'Twin' }, { name: '***' }, { name: 'Başlık Şık' }]));
  builder.text('x\r');
  builder.papx(0, 2, 4);
  const model = builder.read();
  assert.deepEqual([...model.styles.styles.keys()], ['Normal', 'Twin', 'Twin_2', 'Style3', 'BaşlıkŞık']);
  assert.equal(paragraphs(model.body.blocks)[0].props.style, 'BaşlıkŞık');
});

test('a paragraph that names a style the file does not have falls back to the default', () => {
  const builder = styled();
  builder.text('x\r');
  builder.papx(0, 2, 77);
  const model = builder.read();
  const [paragraph] = paragraphs(model.body.blocks);
  assert.equal(paragraph.props.style, undefined);
  assert.equal(resolveRunProps(model, paragraph.props, {}).size, 24);
});

test('the font table: names, alternative names and kinds', () => {
  const builder = new DocBuilder();
  builder.table(FC.SttbfFfn, fontTable(['Times New Roman', 'Arial']));
  builder.text('x\r');
  builder.papx(0, 2, 0);
  const { fonts } = builder.read();
  assert.deepEqual([...fonts.keys()], ['Times New Roman', 'Arial']);
  assert.equal(fonts.get('Arial')!.family, 'swiss');
  assert.equal(fonts.get('Arial')!.pitch, 'fixed');
  assert.equal(fonts.get('Arial')!.panose, '00000000000000000000');
});

test('tracked changes: deleted and inserted runs carry their revision', () => {
  const builder = new DocBuilder();
  const text = 'kept gone new\r';
  builder.text(text);
  builder.papxAll(text);
  builder.table(FC.SttbfRMark, sttb(['Unknown', 'Ayşe']));
  // 2020-03-04 05:06.
  const date = 6 | (5 << 6) | (4 << 11) | (3 << 16) | (120 << 20);
  builder.chpx(5, 10, sprm(0x0800, 1), sprm(0x4863, u16le(1)), sprm(0x6864, u32le(date)));
  builder.chpx(10, 13, sprm(0x0801, 1), sprm(0x4804, u16le(1)));
  const [paragraph] = paragraphs(builder.read().body.blocks);
  const [kept, gone, added] = runsOf(paragraph);
  assert.equal(kept.wrap, undefined);
  assert.deepEqual(findWrapper(gone, 'del'), { kind: 'del', author: 'Ayşe', date: '2020-03-04T05:06:00Z' });
  assert.deepEqual(gone.content, [{ type: 'deletedText', text: 'gone ' }]);
  assert.deepEqual(findWrapper(added, 'ins'), { kind: 'ins', author: 'Ayşe' });
  // Deleted text is not part of the text that shows.
  assert.equal(paragraphText(paragraph), 'kept new');
});

test('a symbol character names its font and character', () => {
  const builder = withRun('(\r', 1, sprm(0x6a09, u16le(1), u16le(0xf0b7)));
  builder.table(FC.SttbfFfn, fontTable(['Times New Roman', 'Symbol']));
  const [paragraph] = paragraphs(builder.read().body.blocks);
  assert.deepEqual(runsOf(paragraph)[0].content, [{ type: 'symbol', char: '', font: 'Symbol' }]);
});

test('formatting pages that point outside themselves or the stream are left out with a warning', () => {
  const builder = withRun('some text\r', 4, sprm(0x0835, 1));
  const model = builder.read({
    patch: ({ word, table, fcLcb }) => {
      // The place of the list of the first character run: beyond the end of its page.
      const [bte] = fcLcb(FC.PlcfBteChpx);
      const page = new DataView(table.buffer, table.byteOffset).getUint32(bte + 8, true) * 512;
      const runs = word[page + 511];
      word[page + (runs + 1) * 4] = 255;
      word[page + 510] = 200;
      // And a paragraph page that is not in the stream at all.
      const [papx] = fcLcb(FC.PlcfBtePapx);
      new DataView(table.buffer, table.byteOffset).setUint32(papx + 8, 0x3fffff, true);
    },
  });
  assert.deepEqual(textOf(model.body.blocks), ['some text']);
  assert.match(model.warnings.join('\n'), /formatting of the document could not be read/);
});
